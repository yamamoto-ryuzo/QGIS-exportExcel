# CHANGELOG — geo_excel

## 1.8.0

- 変数(@…)および詳細フィルター式対応(upstream issue #1)。
  各execスクリプトに `make_expression_context`/`excel_value_text`/
  `evaluate_attach_text` を追加し、式評価を
  グローバル+プロジェクト+レイヤスコープ+地物を含む
  QgsExpressionContext で行うように変更。
  - レイヤ変数内の式(`xlsout_*`)で `@layer_id`・`@project_folder`
    などの @変数が解決されるようになった
    (exportSingle: `variable_based_on_feature`、
    exportSingleBulk: `variable_based_on_feature`/
    `get_exp_variable_for_feature`/`get_variable_for_feature`、
    exportList: `get_layer_variable_evaluated`)。
  - `##Attach::<指定>`・`##ListInsert::<指定>` がフィールド名に
    一致しない場合、属性テーブルの詳細フィルター(式)と同等の
    QGIS式として評価されるようになった
    (例: `"field_a" || '-' || "field_b"`, `@layer_name`)。
    フィールド名指定時の従来動作(QDate/QTime/QDateTime変換含む)は維持。

## 1.7.1

- execスクリプト(exportList/exportSingle/exportSingleBulk)の
  メッセージバー・QMessageBox・QProgressDialogの日本語直書きを
  `QCoreApplication.translate("geo_excel", ...)` へ移行(英語ソース化)。
  テンプレート生成失敗の警告も同様に翻訳化。
- `geo_excel` context に新規24ソースを全9言語で登録。

## 1.7.0

- **翻訳方式を `i18n/<lang>.json` に移行**(Harness横断の方針変更)。
  `.ts`/`.qm`・lupdate/lrelease・`generate_qm.py`・`TRANSLATION.md` を廃止し、
  `{"context": {"英語ソース": "翻訳"}}` のJSONカタログを新設。
  `i18n_json.py`(全プラグイン共通ローダ)の `install_translator()` で
  起動時ロード。`tr()`/`translate()`/`.ui` 文字列は無改変のまま動作。
  未登録の文字列は英語にフォールバック。

## 1.6.0

- **設定ダイアログを新設**(`settings_dialog.py`、メニュー `Settings...`)。
  - 環境チェックの列挙: pywin32(win32com) / Excel COM登録(レジストリ+
    「起動テスト」で実起動確認) / ドキュメントフォルダ / 出力先
  - 既定値(QSettings `geo_excel/*`): 出力先フォルダ・DPI・
    一覧/単票テンプレート指定・自動適用の有無
  - アクティブレイヤの `xlsout_*`/`dpi` レイヤ変数を表で確認・編集・
    「不足変数へ既定値を適用」「自動テンプレートを再生成」
  - 「必要なもの」セクションで実行要件(Excel・pywin32・テンプレート
    プレースホルダ・マップテーマ等)を列挙
- **既定値モジュール `settings.py` を新設**:
  - 出力先既定 = `<ドキュメント>/geo_excel`
    (`QStandardPaths.DocumentsLocation`。マイドキュメント汎用化)
  - `xlsout_output_path_variable` 既定 = `concat('<レイヤ名>_', $id, '.xlsx')`
    (一覧は `<レイヤ名>_list.xlsx`)
  - テンプレート未指定時は `<出力先>/templates/<レイヤ名>_<種別>.xlsx` に
    `##Attach::`/`##ListInsert::` プレースホルダ入りの最小xlsxを
    レイヤのフィールドから自動生成(zipfileによる直接生成)
- **実行時の自動適用**: `_exec_script` が実行前に
  `ensure_layer_variables` で不足している必須レイヤ変数を既定値で補完
  (出力先フォルダ自動作成含む)。QSettings `geo_excel/auto_apply` で
  無効化可(既定ON)。これによりレイヤ変数未設定でも即実行できる
- **bridge メソッド追加**: `settings()`(既定値・環境・要件列挙) /
  `applyDefaults(layer_id, script)` / `checkExcel()`(実起動検証) /
  `openSettings()`
- i18n: 新規UI文字列を9言語の `.ts`/`.qm` に追加

## 1.5.0

- 採用言語セット(AGENTS.md)に合わせ de/es/fr/hi/it/pt/ru/zh の
  翻訳カタログ(`.ts`/`.qm`)を追加。ja と合わせ全9言語を同梱。

## 1.4.0

- 国際化(i18n)対応: UI文字列を英語ソースに統一し、
  `QCoreApplication.translate` で抽出可能な翻訳呼出に変更。
  `QTranslator` による起動時ロードを追加。
- `i18n` に日本語カタログを新設し `ja.qm` をコンパイル。
- `README.md` を英語版に刷新し、日本語ドキュメントを `README_JP.md` に分離。

## 1.3.0

- `metadata.txt` の `qgisMinimumVersion` を 3.40 → 4.0 に更新
  (QGIS4以降専用。`qgisMaximumVersion=4.99`・`supportsQt6`は継続)
- bridge の自動応答を Qt6 の `QMessageBox.StandardButton` スコープenumに修正

## 1.2.0

- **AI/外部制御用ブリッジ**(`bridge.py` の `GeoExcelBridge`)を追加し
  `plugin.bridge` に配線。`qgis_live_driver` の `bridge_call`
  (ファイルJSONL / REST /api/cmd / MCP `qgis_bridge_call` /
  AIアシスタントパネル)と `describe_plugins` カタログから利用可能。
  メソッド(JSON文字列返却):
  - `pluginInfo()` / `status()` / `listScripts()`
  - `layerVariables(layer_id)` / `setLayerVariables(vars, layer_id)`
    — 帳票出力に必要なレイヤ変数 `xlsout_*`・`dpi` の取得/設定
  - `run(script, feature_id, confirm)` — 帳票実行。scriptは
    list|single|bulk|ファイル名|index。スクリプト内のmodal
    QMessageBox(上書き/続行確認・完了通知)は実行中だけ自動応答に
    差し替え、confirm=False(既定)では確認系をCancel=中断として扱う
    安全側の既定。messageBar().pushCritical の呼出も収集して返す。
    Excel起動・生成中はGUIスレッドをブロックする(メニュー実行と同じ)
- `plugin.py` の `_exec_script` がexec名前空間を返すように拡張
  (スクリプト側の `excel_result`/`output_path` をbridgeが取得するため)

## 1.1.0

- プラグインID・フォルダ名を `export_excel` から `geo_excel` に改名
  (他プラグインと命名規則を統一)。deploy先ディレクトリ・メニュー名・
  upstream gitdir(`.upstream/QGIS-exportExcel.git`)の core.worktree も追従。
  upstreamリポジトリ名(QGIS-exportExcel)自体は変更なし

## 1.0.0

- QGIS-exportExcel (https://github.com/yamamoto-ryuzo/QGIS-exportExcel) の
  レイヤアクション用スクリプト群(exportList/exportSingle/exportSingleBulk)を
  無改変で同梱し、プラグインメニューから実行できる簡易ラッパーとして組み込み。
  - `[% @layer_id %]` はアクティブなベクターレイヤ、`[% $id %]` は選択地物で
    実行時に置換する
  - スクリプト内の `PyQt5` 固定importは exec 直前にソースを `qgis.PyQt` へ
    書き換えて吸収(Qt6対応。QGIS4のimportフックがPyQt5を禁止するため
    sys.modules shim は不可)。`Qt.WaitCursor` も `Qt.CursorShape.WaitCursor`
    へ同時に書き換える(Qt6のスコープenum対応)
- 実行には win32com(pywin32) と Excel が必要(upstreamと同じ前提)
