# NotebookML_MasterDoc

Google Drive 上のフォルダからマスターDocを生成するスクリプトです。

## 出力モード
`NotebookML_MasterDoc.gs` 冒頭の設定で出力先フォルダ構成を選べます。

```javascript
const DEST_MODE = 'hierarchical'; // 'flat' で平置き、'hierarchical' で階層構造
const DEST_MAX_DEPTH = 3;         // 階層構造にする際の最大深さ (0=制限なし)
```

`DEST_MODE` を `flat` にすると全てのマスターDocが直下に作成され、
`hierarchical` にすると元フォルダの階層を最大 `DEST_MAX_DEPTH` 階層まで再現します。

