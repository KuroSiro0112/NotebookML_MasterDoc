/***** 設定（ID でも URL でもOK）*********************************************/
// 走査の起点（この配下の全フォルダを処理）
const ROOT_FOLDER_INPUT = 'https://drive.google.com/drive/folders/1jbHEIottL4zAnRKxtQY5_I4ZedX_F20F';
// 生成/更新したマスターDocの保管先
const DEST_FOLDER_INPUT  = 'https://drive.google.com/drive/folders/1WjqZqfeII7lY1Bmndl2ZcvEcrYGaQ9Tl';

// 走査
const RECURSIVE = true;              // サブフォルダも処理
// 出力オプション
const INSERT_TOC = false;            // マスターDoc先頭に目次を入れる
const PAGE_BREAK_BETWEEN_DOCS = true;// ソースDocの区切りに改ページ
// 文字数制御（ライト構成：基本無制限、必要なら制限）
const MAX_CHARS_PER_SOURCE_DOC = null; // 1ソースDoc取り込み上限(null=制限なし)
const MAX_TOTAL_CHARS_PER_MASTER = null; // マスターDoc合計上限(null=制限なし)
// 実行時間制御
const TIME_BUDGET_MS    = 24 * 60 * 1000; // 24分（30分制限の手前）
const SAFETY_MARGIN_MS  = 60 * 1000;      // 1分の余裕を見て切り上げ
const RESUME_DELAY_MIN  = 1;              // 次の継続まで1分
const RESUME_HANDLER    = 'runPendingWorker__resume';
const SINGLE_RUNNER_FN  = 'buildOrUpdateMasterForSingleFolder'; // フォルダ1件処理関数名

/***** シート名 ***************************************************************/
const SHEET_QUEUE = 'Queue';
const SHEET_LOG   = 'RunLog';
const SHEET_META  = 'Meta';

/***** 便利ユーティリティ *****************************************************/
function formatDate_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'); }
function formatDuration_(ms){ const s=Math.floor(ms/1000); const h=String(Math.floor(s/3600)).padStart(2,'0'); const m=String(Math.floor((s%3600)/60)).padStart(2,'0'); const ss=String(s%60).padStart(2,'0'); return `${h}:${m}:${ss}`; }
function getIdFromUrlOrId_(u){ const m=String(u).match(/[-\w]{25,}/); return m?m[0]:u; }
function getFolderByUrlOrId_(u){ return DriveApp.getFolderById(getIdFromUrlOrId_(u)); }

function getSS_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function getOrCreateSheet_(name, headers){
  const ss=getSS_(); let sh=ss.getSheetByName(name);
  if(!sh){ sh=ss.insertSheet(name); }
  if(headers && sh.getLastRow()===0){ sh.appendRow(headers); }
  return sh;
}
function writeMeta_(key,value){
  const sh=getOrCreateSheet_(SHEET_META,['key','value']);
  const last=Math.max(sh.getLastRow(),1);
  const keys=sh.getRange(1,1,last,1).getValues().flat();
  let r=keys.indexOf(key)+1;
  if(r<=0){ r=last+1; sh.appendRow([key,value]); }
  else     { sh.getRange(r,2).setValue(value); }
}
function readMeta_(key){
  const sh=getOrCreateSheet_(SHEET_META,['key','value']);
  const last=Math.max(sh.getLastRow(),1);
  const keys=sh.getRange(1,1,last,1).getValues().flat();
  const r=keys.indexOf(key)+1;
  return (r>0)? String(sh.getRange(r,2).getValue()): '';
}

function deleteTimeTriggersByHandler_(handler){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()===handler) ScriptApp.deleteTrigger(t);
  });
}
function scheduleResume_(){
  deleteTimeTriggersByHandler_(RESUME_HANDLER);
  ScriptApp.newTrigger(RESUME_HANDLER).timeBased().after(RESUME_DELAY_MIN*60*1000).create();
}
function clearResume_(){ deleteTimeTriggersByHandler_(RESUME_HANDLER); }

/***** Queue/RunLog 操作 ******************************************************/
function ensureQueueHeaders_(){
  getOrCreateSheet_(SHEET_QUEUE, [
    'folderId','folderPath','folderName','status','note','lastRun','newPartUrls','runStart','runEnd'
  ]);
  getOrCreateSheet_(SHEET_LOG, [
    'フォルダ名','フォルダパス','マスターDoc ID','マスターDoc URL','Part',
    '元Docタイトル','元Doc URL','元Doc ID','取り込み文字数(実測)',
    'フォルダ内インデックス','開始','終了','エラー','状態'
  ]);
  getOrCreateSheet_(SHEET_META, ['key','value']);
}
function appendRunLog_(row){
  const sh=getOrCreateSheet_(SHEET_LOG);
  sh.appendRow(row);
}
function findNextPendingRow_(shQueue){
  const last=shQueue.getLastRow(); if(last<2) return null;
  const vals=shQueue.getRange(2,1,last-1,9).getValues();
  for(let i=0;i<vals.length;i++){
    const status=vals[i][3]; // D
    if(status==='PENDING' || status==='ERROR') return {row:i+2, values:vals[i]};
  }
  return null;
}
function setQueueStatus_(sh,row,{status,note,runStart,runEnd,lastRun}){
  if(status!==undefined)  sh.getRange(row,4).setValue(status);
  if(note!==undefined)    sh.getRange(row,5).setValue(note);
  if(lastRun!==undefined) sh.getRange(row,6).setValue(lastRun);
  if(runStart!==undefined)sh.getRange(row,8).setValue(runStart);
  if(runEnd!==undefined)  sh.getRange(row,9).setValue(runEnd);
}

/***** Queue を ROOT_FOLDER_INPUT から構築 *************************************/
function rebuildQueueFromRoot_(){
  const root=getFolderByUrlOrId_(ROOT_FOLDER_INPUT);
  const sh=getOrCreateSheet_(SHEET_QUEUE);
  sh.clearContents();
  ensureQueueHeaders_();
  let rows=[];
  function walk(f, path){
    const thisPath = path ? `${path}/${f.getName()}` : f.getName();
    rows.push([f.getId(), thisPath, f.getName(), 'PENDING', '', '', '', '', '']);
    if(RECURSIVE){
      const it=f.getFolders();
      while(it.hasNext()){ walk(it.next(), thisPath); }
    }
  }
  walk(root, '');
  if(rows.length>0) sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

/***** マスターDoc ID の永続化（フォルダ単位） **********************************/
function getMasterIdKey_(folderId){ return `MASTER_DOC_${folderId}`; }
function getOrCreateMasterDocForFolder_(folderId, folderName, folderPath){

  // folderId は URL/ID どちらでも受け付ける
  folderId = getIdFromUrlOrId_(folderId);
  if(!folderId) throw new Error('folderId is required');

  const props=PropertiesService.getScriptProperties();
  const key=getMasterIdKey_(folderId);
  const saved=props.getProperty(key);
  let masterId;
  try{
    if(saved){ DocumentApp.openById(saved); masterId = saved; }
  }catch(e){ /* 落ちていたら作り直し */ }

  // 保存先のフォルダ階層を folderPath の第3階層まで作成
  const destRoot=getFolderByUrlOrId_(DEST_FOLDER_INPUT);
  const segments=String(folderPath||'').split('/').slice(0,-1).slice(0,3);
  let destFolder=destRoot;
  for(const name of segments){
    const it=destFolder.getFoldersByName(name);
    destFolder = it.hasNext() ? it.next() : destFolder.createFolder(name);
  }

  if(!masterId){
    const name = `NotebookLM_Master_${folderName}`;
    const doc  = DocumentApp.create(name);
    masterId = doc.getId();
    props.setProperty(key, masterId);
  }

  DriveApp.getFileById(masterId).moveTo(destFolder);
  return masterId;
}

/***** Googleドキュメントの本文を取得（上限考慮） ******************************/
function fetchDocText_(docId){
  const doc = DocumentApp.openById(docId);
  let text = doc.getBody().getText() || '';
  if(MAX_CHARS_PER_SOURCE_DOC && text.length>MAX_CHARS_PER_SOURCE_DOC){
    text = text.slice(0, MAX_CHARS_PER_SOURCE_DOC);
  }
  return {title: doc.getName(), text, url: doc.getUrl()};
}

/***** フォルダ1件を処理（Queue から呼ばれる中核） *****************************/
function buildOrUpdateMasterForSingleFolder(folderId, folderPath, folderName, opt={}){
  // folderId は URL/ID どちらでも受け付ける
  folderId = getIdFromUrlOrId_(folderId);
  if(!folderId) throw new Error('folderId is required');
  const folder = DriveApp.getFolderById(folderId);
  const masterId = getOrCreateMasterDocForFolder_(folderId, folderName, folderPath);
  const master   = DocumentApp.openById(masterId);
  const body     = master.getBody();
  // クリア
  body.clear();
  // 見出し
  body.appendParagraph(folderName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  if(INSERT_TOC){
    body.appendTableOfContents(DocumentApp.ParagraphHeading.HEADING1);
  }

  // ソースDoc列挙（同フォルダ直下のみ・ライト構成）
  const files=[];
  const it = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  while(it.hasNext()){ files.push(it.next()); }
  files.sort((a,b)=> a.getName().localeCompare(b.getName(),'ja'));

  let total = 0;
  let index = 0;
  for(const f of files){
    index++;
    const start = new Date();
    try{
      const {title,text,url} = fetchDocText_(f.getId());
      if(PAGE_BREAK_BETWEEN_DOCS && body.getNumChildren()>0){
        body.appendPageBreak();
      }
      // 見出し + リンク
      const p = body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      p.setLinkUrl(url);
      // 本文
      body.appendParagraph(text);
      total += text.length;

      if(opt.appendLog){
        opt.appendLog([
          folderName, folderPath, masterId, master.getUrl(), '', // Partは未使用
          title, url, f.getId(), text.length,
          index, formatDate_(start), formatDate_(new Date()), '', 'UPDATED'
        ]);
      }
      if(MAX_TOTAL_CHARS_PER_MASTER && total>=MAX_TOTAL_CHARS_PER_MASTER) break;
    }catch(e){
      if(opt.appendLog){
        opt.appendLog([
          folderName, folderPath, masterId, master.getUrl(), '',
          f.getName(), f.getUrl(), f.getId(), '',
          index, formatDate_(start), formatDate_(new Date()),
          String(e).slice(0,500), 'ERROR'
        ]);
      }
    }
  }
  master.saveAndClose();
  return {masterId, masterUrl: master.getUrl(), count: files.length};
}

/***** ランナ（時間予算を見ながらPENDING/ERRORを一掃） *************************/
function buildOrUpdateAllMasterDocs(){
  ensureQueueHeaders_();
  const shQ=getOrCreateSheet_(SHEET_QUEUE);

  // Queue が空なら作る
  if(shQ.getLastRow()<2) rebuildQueueFromRoot_();

  // 実行全体の開始/ログ
  const start=new Date();
  writeMeta_('lastRunStart', formatDate_(start));
  writeMeta_('lastRunNote', 'start');

  // すぐ一塊回す（時間内まで）＋必要なら継続トリガー
  runPendingWorker__resume();
}

function runPendingWorker__resume(){
  ensureQueueHeaders_();
  const startedAt=Date.now();
  const shQ=getOrCreateSheet_(SHEET_QUEUE);

  let processed=0;
  while(true){
    // 時間予算
    if(Date.now()-startedAt > (TIME_BUDGET_MS - SAFETY_MARGIN_MS)){
      if(findNextPendingRow_(shQ)){ scheduleResume_(); writeMeta_('lastRunNote','time-sliced resume scheduled'); }
      break;
    }

    const next=findNextPendingRow_(shQ);
    if(!next){
      // 完了
      clearResume_();
      const end=new Date();
      writeMeta_('lastRunEnd', formatDate_(end));
      writeMeta_('lastRunDuration', formatDuration_(end - new Date(readMeta_('lastRunStart'))));
      break;
    }

    const row=next.row;
    const [folderId, folderPath, folderName] = [ next.values[0], next.values[1], next.values[2] ];
    setQueueStatus_(shQ,row,{status:'RUNNING',note:'',runStart:formatDate_(new Date()),runEnd:''});

    const startOne=new Date();
    try{
      const result = this[SINGLE_RUNNER_FN](folderId, folderPath, folderName, {appendLog: appendRunLog_});
      setQueueStatus_(shQ,row,{
        status:'UPDATED',
        note:`OK (${result.count} docs)`,
        lastRun:formatDate_(new Date()),
        runEnd:formatDate_(new Date())
      });
    }catch(e){
      setQueueStatus_(shQ,row,{
        status:'ERROR',
        note:String(e).slice(0,500),
        lastRun:formatDate_(new Date()),
        runEnd:formatDate_(new Date())
      });
      // エラーでも続行（他の行は動かす）
    }
    processed++;
  }
  writeMeta_('lastChunk', `${processed} rows @ ${formatDate_(new Date())}`);
}

/***** 手動停止 ***************************************************************/
function cancelRun(){ clearResume_(); }

/***** onOpen（メニュー）******************************************************/
function onOpen(){
  SpreadsheetApp.getUi().createMenu('MasterDoc')
    .addItem('PENDING を全部実行','buildOrUpdateAllMasterDocs')
    .addItem('Queue を作り直す','rebuildQueueFromRoot_')
    .addSeparator()
    .addItem('継続を停止','cancelRun')
    .addToUi();
}
