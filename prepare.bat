SET /P ANSWER="Do you want to do that install libraries and make env file? (Y/N)"
if /i {%ANSWER%}=={y} (goto :exec)
EXIT

:exec
pip install openpyxl
pip install python-dotenv
pip install sendgrid
pip install pyinstaller
copy env-example.txt .env
notepad .env