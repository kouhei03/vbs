# ファイルシステム・ストレージ自動管理 (No. 1-50)

Windows Script Host用のVBScript例です。Windowsのコマンドプロンプトから、
各ファイルに記載されたUsageに従って個別に実行してください。

```text
cscript //nologo 001_ListFilesRecursive.vbs <folder> <log-file>
cscript //nologo 003_DeleteOldFiles.vbs <folder> <days>
cscript //nologo 003_DeleteOldFiles.vbs <folder> <days> DELETE
cscript //nologo 018_ZipFolder.vbs <folder> <new-archive.zip>
```

| No. | スクリプト | 内容 |
|---:|---|---|
| 1 | `001_ListFilesRecursive.vbs` | 再帰ファイル一覧とサイズをログ出力 |
| 2 | `002_SortByExtension.vbs` | 拡張子別フォルダーへの振り分け |
| 3 | `003_DeleteOldFiles.vbs` | 古いファイルの表示／明示指定時の削除 |
| 4 | `004_SplitJoinText.vbs` | テキストの行単位分割／結合 |
| 5 | `005_ReplaceText.vbs` | テキスト置換と元ファイルの`.bak`保存 |
| 6 | `006_FilterCsvColumn.vbs` | 指定列の値によるCSV行抽出 |
| 7 | `007_ConcatenateTextFiles.vbs` | 同一拡張子のテキストを名前順に連結 |
| 8 | `008_SetFileTimestamp.vbs` | ファイル更新日時の変更 |
| 9 | `009_FileHash.vbs` | `certutil`によるSHA-256計算 |
| 10 | `010_CreateFolderTree.vbs` | 親階層を含むフォルダー作成 |
| 11 | `011_BestEffortFileErase.vbs` | 単一ファイルのベストエフォート上書き・削除 |
| 12 | `012_FolderSizeReport.vbs` | 再帰的なファイル数・合計サイズ集計と任意の閾値警告 |
| 13 | `013_CopyMoveNoClobber.vbs` | 衝突を避けるコピー／移動 |
| 14 | `014_BulkFileAttributes.vbs` | 直下ファイル属性の付与／解除 |
| 15 | `015_AddNameAffixes.vbs` | ファイル名への接頭辞・接尾辞追加 |
| 16 | `016_NormalizeFileNames.vbs` | 全角ASCII・空白を半角化し、空白をアンダースコアへ置換 |
| 17 | `017_NumberFiles.vbs` | 名前順の連番リネーム |
| 18 | `018_ZipFolder.vbs` | Windows Shellによる新規ZIP作成 |
| 19 | `019_MirrorCopy.vbs` | 更新ファイルの再帰コピー（宛先の余分なファイルは保持） |
| 20 | `020_CreateShortcut.vbs` | `.lnk`ショートカット作成 |
| 21 | `021_ReadIniValue.vbs` | INI形式のセクション・キー読み取り |
| 22 | `022_ScriptFolderPath.vbs` | 実行中スクリプトの絶対パス表示 |
| 23 | `023_PatchBinaryBytes.vbs` | バックアップ付きバイナリ範囲置換 |
| 24 | `024_ConvertTextEncoding.vbs` | BOM判定または指定文字コードからUTF-8へ変換 |
| 25 | `025_RotateLog.vbs` | ログ退避と新規ログ作成 |
| 26 | `026_ReadFolderAcl.vbs` | `icacls`によるACL照会 |
| 27 | `027_CreateTempFolder.vbs` | TEMP配下の一意な作業フォルダー作成 |
| 28 | `028_NormalizeImageExtensions.vbs` | 画像ヘッダーに基づく拡張子正規化 |
| 29 | `029_EmptyRecycleBin.vbs` | ごみ箱の内容表示／確認語指定時の削除要求 |
| 30 | `030_WaitForFileUnlock.vbs` | ファイルが利用可能になるまでの期限付き再試行 |
| 31 | `031_CheckDriveFreeSpace.vbs` | ドライブ空き容量と閾値の確認 |
| 32 | `032_FindLongPaths.vbs` | 指定長を超えるパスの検出 |
| 33 | `033_IncrementalCopy.vbs` | 更新日時・サイズに基づく差分コピー |
| 34 | `034_ReadFileVersion.vbs` | ファイルバージョン情報表示 |
| 35 | `035_AppendTimestampLog.vbs` | 日時付きメッセージのログ追記 |
| 36 | `036_ExpandFilenamePlaceholders.vbs` | `{DATE}`・`{USER}`のファイル名展開 |
| 37 | `037_NormalizeCsvLines.vbs` | CSVの末尾空白除去と改行統一 |
| 38 | `038_RetainBackupGenerations.vbs` | バックアップ世代作成と保持数管理 |
| 39 | `039_WaitForLockFile.vbs` | ロックファイル消失の期限付き待機 |
| 40 | `040_RetryNetworkCopy.vbs` | 指定先へのコピー再試行 |
| 41 | `041_ExtractErrorLines.vbs` | キーワードを含むログ行の抽出 |
| 42 | `042_RemoveEmptyFolders.vbs` | 空フォルダー検出／指定時の削除 |
| 43 | `043_ListSpecialFolders.vbs` | WSH特殊フォルダーのパス表示 |
| 44 | `044_WaitForDownloadStable.vbs` | ダウンロードファイルのサイズ安定待機 |
| 45 | `045_CheckTextLength.vbs` | テキスト文字数と最小値の検証 |
| 46 | `046_CompareCsvHeaders.vbs` | 複数CSVのヘッダー比較 |
| 47 | `047_ReadKeyValueConfig.vbs` | `key=value`設定ファイルの読み取り |
| 48 | `048_ProtectWhitelistFiles.vbs` | 保護対象を除外した削除候補表示（削除なし） |
| 49 | `049_CreateUserFoldersAcl.vbs` | ユーザーフォルダーとModify ACLの作成 |
| 50 | `050_CreateDailyFolder.vbs` | 日付名フォルダー作成 |

## 注意事項

- `APPLY`、`DELETE`、確認語などの指定がある処理は、その実行語を付けずに対象一覧・プレビューを確認してください。対象のバックアップも含め、必要なデータを事前に保護してください。
- No. 11は媒体の特性、スナップショット、バックアップ等により復元不能を保証できません。機密データの安全な消去用途には使わないでください。
- No. 6は引用符内のカンマや複数行フィールドを扱う完全なCSVパーサーではありません。No. 24はBOMなしファイルの文字コードを自動判定できないため、文字コードを指定してください。
- ファイル操作の一部は指定フォルダー直下のみを対象にします。各スクリプトの冒頭コメントとUsageを確認してください。
- ACL変更、属性変更、削除、ごみ箱操作などは、対象環境と権限を確認したうえで実行してください。管理者権限やユーザーアカウント情報を自動取得・昇格する処理はありません。
