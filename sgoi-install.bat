SET /P ANSER="ライブラリをインストールします。よろしいですか？ (Y/N)"
if /i {%ANSWER%}=={y} (goto :install-library)
if /i {%ANSWER%}=={yes} (goto :install-library)
SET /P ANSER="環境変数ファイルを作成（既にある場合は上書き）しますか？ (Y/N)"
if /i {%ANSWER%}=={y} (goto :make-env)
if /i {%ANSWER%}=={yes} (goto :make-env)
SET /P ANSER="環境変数ファイルを編集しますか？ (Y/N)"
if /i {%ANSWER%}=={y} (goto :edit-env)
if /i {%ANSWER%}=={yes} (goto :edit-env)
EXIT

:install-library
pip install openpyxl
pip install python-dotenv
pip install sendgrid
pip install pyinstaller

:make-env
copy env-example.txt .env

:edit-env
notepad .env