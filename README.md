# SGOI😲
SendGrid API を使用してメールを一斉送信するソフトウェア。

## 導入方法 (Windows)
1. [Microsoft Store から Python をインストール](https://www.microsoft.com/store/productId/9PJPW5LDXLZ5)する。
2. 「prepare.bat」を実行して必要なライブラリのインストールと、API キーの設定を行う。
※ セキュリティで弾かれる場合、右クリックの「プロパティ」を開き、「全般」タブにあるセキュリティの「許可する」にチェックを入れる。
3. 「sgoi.py」を実行する。

## 使用方法
文章ファイルと変数ファイルを用意し、「sgoi.py」でぞれぞれのファイルを指定してメールを送信します。
### 文章ファイル
送信する文章が記載されたテキストファイルです。可変部分は {変数名} で記載できます。
「txt-example.txt」を複製して使用してください。
### 変数ファイル
下記の必須項目に加え、独自の各変数を指定する Excel ブックファイルです。
「val-example.xlsx」を複製して使用してください。
※ sgoi_html_content は独自の変数名として使用することはできません。
#### 必須変数
- sgoi_from_email : メール送信元（From）のメールアドレス
- sgoi_to_emails : メール送信先（To）のメールアドレス
- sgoi_subject : メールの題名

## バージョン履歴
- v1.0.1 (2022-07-05)
  - 送信待機間隔を 1 秒から 0.2 秒に変更
- v1.0 (2022-07-04)
  - GUI による基本的な送信機能のみ実装


## 必要なライブラリ
- openpyxl
- python-dotenv
- sendgrid
