# SGOI - SendGrid Operation Interface

from dotenv import load_dotenv
import logging
import openpyxl
import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
import threading
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox

class sgoiMail():

    @staticmethod
    def hello():
        return '''
        ///===---
        SendGrid Operation Interface
        v1.0.2 @fts141     ---===///
        '''

    def verify(self):

        def checkFiles():

            # 文章ファイルの確認
            try:
                self.txtFileObj = open(self.entry_txtFile.get(), 'r', encoding='utf-8')
                self.txtFile = self.txtFileObj.read()
            except Exception as e:
                messagebox.showerror('文章ファイルエラー', '文章ファイルが読み込めませんでした。\n再度ファイルを選択してください。\n\n{}'.format(e))
                return False

            # 変数ファイルの確認
            try:
                self.valFileWb = openpyxl.load_workbook(self.entry_valFile.get())
                self.valFileWs = self.valFileWb.active
                self.valKeys = tuple(self.valFileWs.rows)[0]
            except Exception as e:
                messagebox.showerror('変数ファイルエラー', '変数ファイルが読み込めませんでした。\n再度ファイルを選択してください。\n\n{}'.format(e))
                return False

            # 変数ファイルの項目確認
            self.reqKeys = ('sgoi_from_email', 'sgoi_to_emails', 'sgoi_subject')
            self.denyKey = 'sgoi_html_content'
            cnt = 0; f = False; f0 = False; f1 = False; f2 = False
            for key in self.valKeys:
                if key.value == self.denyKey: f = True
                elif key.value == self.reqKeys[0] and not f0: f0 = True; cnt += 1
                elif key.value == self.reqKeys[1] and not f1: f1 = True; cnt += 1
                elif key.value == self.reqKeys[2] and not f2: f2 = True; cnt += 1
            if f:
                messagebox.showerror('変数ファイルエラー', '変数ファイルに禁止項目があります。\nファイルを確認してください。\n\n{}'.format(self.denyKey))
                return False
            elif not(f0 and f1 and f2):
                messagebox.showerror('変数ファイルエラー', '変数ファイルに必須項目がありません。\nファイルを確認してください。\n\n{}'.format(self.reqKeys))
                return False
            elif cnt != 3:
                messagebox.showerror('変数ファイルエラー', '変数ファイルの必須項目が重複しています。\nファイルを確認してください。\n\n{}'.format(self.reqKeys))
                return False

            return True

        def prepare():
            
            # 変数ファイルの1行を辞書型で格納し、リストにまとめる
            self.emails = []
            for rCnt, row in enumerate(self.valFileWs.rows):

                if rCnt == 0: continue
                valDict = {}
                for cnt, tpl in enumerate(self.valKeys):
                    valDict['{}'.format(tpl.value)] = row[cnt].value

                # 文章ファイルから文字を置き換え、辞書に格納
                txt = self.txtFile
                for key in valDict.keys():
                    if key not in self.reqKeys:
                        repStr = r'{' + '{}'.format(key) + r'}'
                        txt = txt.replace(repStr, '{}'.format(valDict[key]))
                valDict[self.denyKey] = txt

                self.emails.append(valDict)

        if checkFiles(): prepare()
        else: return False

        self.previewIdx = 0
        self.showPreview(self.previewIdx)
        return True

    def send(self):

        self.cancel = False
        self.progBar.configure(maximum=len(self.emails))
        sg = SendGridAPIClient(os.getenv('SENDGRID_API_KEY'))

        self.insertActivity('{} メール送信処理を開始しました。\n'.format(time.strftime('%Y/%m/%d %H:%M:%S')))

        for cnt, email in enumerate(self.emails):
            self.progbarVal.set(cnt + 1)
            self.showPreview(cnt)
            self.insertActivity('{} #{}, {} -> {}\n'.format(time.strftime('%Y/%m/%d %H:%M:%S'), cnt, email['sgoi_from_email'], email['sgoi_to_emails']))
            try:
                message = {
                    'from_email': email['sgoi_from_email'],
                    'to_emails': email['sgoi_to_emails'],
                    'subject': email['sgoi_subject'],
                    'html_content': email['sgoi_html_content'].replace('\n', '<br>')
                    }
                response = sg.send(Mail(**message)) 
                self.insertActivity(' [OK] 正常に送信されました。({})\n'.format(response.status_code))

                # logging.info('メール送信：{} -> responseCode:{}'.format(self.to_emails, response.status_code))
                # logging.debug('== headers ==\n{}\n== body ==\n{}\n'.format(response.headers, response.body),)
            except Exception as e:
                self.insertActivity('![NG] エラーが発生しました。\n')
                # pass

                # logging.warning('メール送信に失敗：{}'.format(e))
            time.sleep(0.2)
            if self.cancel: break

        if self.cancel:
            self.insertActivity('{} メール送信処理を中止しました。\n'.format(time.strftime('%Y/%m/%d %H:%M:%S')))
            messagebox.showwarning('中止', '{} 件の送信をし、処理を中止しました。\n詳細はアクティビティを確認してください。'.format(cnt + 1))
        else:
            self.insertActivity('{} メール送信処理が完了しました。\n'.format(time.strftime('%Y/%m/%d %H:%M:%S')))
            messagebox.showinfo('完了', '{} 件の送信処理が完了しました。'.format(cnt + 1))
        self.enableWidgets('initial')

    def enableWidgets(self, mode):

        if mode == 'initial':
            self.label_guideVar.set('文章ファイルおよび変数ファイルを選択し、検証してください。')
            
            self.label_txtFile.configure(state='enable')
            self.entry_txtFile.configure(state='enable')
            self.button_txtFile.configure(state='enable')
            self.label_valFile.configure(state='enable')
            self.entry_valFile.configure(state='enable')
            self.button_valFile.configure(state='enable')
            self.label_arrow.configure(state='enable')
            self.button_verify.configure(state='enable')

            self.button_prev.configure(state='disable')
            self.button_next.configure(state='disable')
            self.button_again.configure(state='disable')
            self.button_exec.configure(state='disable')
            self.button_cancel.configure(state='disable')
            
            self.showPreview(-1)
            self.label_idxVar.set('ヽ(´･д･｀)ﾉ')
            self.updateProgBar(0)

        elif mode == 'verified':
            self.label_guideVar.set('内容を確認し、メールの送信開始をするか、ファイルを再選択してください。')

            self.label_txtFile.configure(state='disable')
            self.entry_txtFile.configure(state='disable')
            self.button_txtFile.configure(state='disable')
            self.label_valFile.configure(state='disable')
            self.entry_valFile.configure(state='disable')
            self.button_valFile.configure(state='disable')
            self.label_arrow.configure(state='disable')
            self.button_verify.configure(state='disable')

            self.preview.config(state='disable')
            self.button_again.configure(state='enable')
            self.button_exec.configure(state='enable')

            self.button_cancel.configure(state='disable')
            self.progbarVal.set(0)

        elif mode == 'started':
            self.label_guideVar.set('メールを送信しています...')

            self.label_txtFile.configure(state='disable')
            self.entry_txtFile.configure(state='disable')
            self.button_txtFile.configure(state='disable')
            self.label_valFile.configure(state='disable')
            self.entry_valFile.configure(state='disable')
            self.button_valFile.configure(state='disable')
            self.label_arrow.configure(state='disable')
            self.button_verify.configure(state='disable')

            # self.notebook.configure(state='enable')
            self.preview.config(state='disable')
            self.button_prev.configure(state='disable')
            self.button_next.configure(state='disable')
            self.button_again.configure(state='disable')
            self.button_exec.configure(state='disable')

            self.button_cancel.configure(state='enable')
            self.updateProgBar(0)

    def insertActivity(self, val):
        self.activity.configure(state='normal')
        self.activity.insert(tk.END, val)
        self.activity.configure(state='disabled')

    def showPreview(self, idx=-1):

        self.preview.configure(state='normal')
        self.button_prev.configure(state='disable')
        self.button_next.configure(state='disable')
        self.preview.delete('1.0', tk.END)
        if idx >= 0:
            email = self.emails[idx]
            if idx != 0: self.button_prev.configure(state='enable')
            if idx != len(self.emails) - 1: self.button_next.configure(state='enable')
            self.preview.insert(tk.END, '[ML FROM] {}\n'.format(email['sgoi_from_email']))
            self.preview.insert(tk.END, '[MAIL TO] {}\n'.format(email['sgoi_to_emails']))
            self.preview.insert(tk.END, '[SUBJECT] {}\n'.format(email['sgoi_subject']))
            self.preview.insert(tk.END, '\n{}'.format(email['sgoi_html_content']))
            self.label_idxVar.set('{} / {} 件目を表示中'.format(idx + 1, len(self.emails)))
        self.preview.configure(state='disabled')
      
    def updateProgBar(self, val):
        self.progbarVal.set(val)

    def txtFile_clicked(self):
        path = filedialog.askopenfilename(filetypes=[('文章ファイル','*.txt')])
        self.entry_txtFile.delete(0, tk.END)
        self.entry_txtFile.insert(tk.END, path)

    def valFile_clicked(self):
        path = filedialog.askopenfilename(filetypes=[('変数ファイル','*.xlsx; *.xls')])
        self.entry_valFile.delete(0, tk.END)
        self.entry_valFile.insert(tk.END, path)

    def verify_clicked(self):
        if self.verify():
            self.enableWidgets('verified')

    def prev_clicked(self):
        self.previewIdx -= 1
        self.showPreview(self.previewIdx)

    def next_clicked(self):
        self.previewIdx += 1
        self.showPreview(self.previewIdx)

    def again_clicked(self):
        self.enableWidgets('initial')

    def start_clicked(self):
        if not messagebox.askyesno('確認','メールの送信を開始してもよろしいですか？'): return
        self.enableWidgets('started')
        self.thread = threading.Thread(target=self.send)
        self.thread.start()

    def cancel_clicked(self):
        if messagebox.askyesno('送信を中止','メールの送信を中止してもよろしいですか？'): self.cancel = True

    def __init__(self):   
        self.root = tk.Tk()

        # ウィンドウ
        self.root.title('SendGrid Operation Interface')
        self.root.geometry('1000x700+50+50')
        self.root.resizable(width=False, height=False)
        self.root.attributes('-alpha', 0.95)

        # ファイル選択
        self.label_txtFile = ttk.Label(self.root, text='文章ファイル', anchor=tk.CENTER)
        self.label_txtFile.place(x=0, y=0, width=75, height=25)
        self.entry_txtFile = ttk.Entry(self.root)
        self.entry_txtFile.place(x=75, y=0, width=375, height=25)
        self.button_txtFile = ttk.Button(self.root, text='選択', command=self.txtFile_clicked)
        self.button_txtFile.place(x=450, y=0, width=50, height=25)

        self.label_valFile = ttk.Label(self.root, text='変数ファイル', anchor=tk.CENTER)
        self.label_valFile.place(x=0, y=25, width=75, height=25)
        self.entry_valFile = ttk.Entry(self.root)
        self.entry_valFile.place(x=75, y=25, width=375, height=25)
        self.button_valFile = ttk.Button(self.root, text='選択', command=self.valFile_clicked)
        self.button_valFile.place(x=450, y=25, width=50, height=25)

        self.label_arrow = ttk.Label(self.root, text='▶', anchor=tk.CENTER)
        self.label_arrow.place(x=500, y=0, width=25, height=50)

        self.button_verify = ttk.Button(self.root, text='検証', command=self.verify_clicked)
        self.button_verify.place(x=525, y=0, width=75, height=50)

        # プレビュー
        self.preview = scrolledtext.ScrolledText(self.root)
        self.preview.place(x=0, y=50, width=600, height=600)

        self.button_prev = ttk.Button(self.root, text='<', command=self.prev_clicked)
        self.button_prev.place(x=0, y=650, width=50, height=25)
        self.button_next = ttk.Button(self.root, text='>', command=self.next_clicked)
        self.button_next.place(x=50, y=650, width=50, height=25)

        self.label_idxVar = tk.StringVar()
        self.label_idx = ttk.Label(self.root, textvariable=self.label_idxVar, anchor=tk.CENTER)
        self.label_idx.place(x=100, y=650, width=160, height=25)

        # ボタン群
        self.button_again = ttk.Button(self.root, text='ファイルを再選択', command=self.again_clicked)
        self.button_again.place(x=260, y=650, width=120, height=25)
        self.button_exec = ttk.Button(self.root, text='メールの送信開始', command=self.start_clicked)
        self.button_exec.place(x=380, y=650, width=120, height=25)
        self.button_cancel = ttk.Button(self.root, text='送信を中止', command=self.cancel_clicked)
        self.button_cancel.place(x=510, y=650, width=90, height=25)
        
        # ガイド
        self.label_guideVar = tk.StringVar()
        self.label_guide = ttk.Label(self.root, textvariable=self.label_guideVar)
        self.label_guide.place(x=0, y=675, width=600, height=25)

        # アクティビティ
        self.activity = scrolledtext.ScrolledText(self.root)
        self.activity.place(x=600, y=0, width=400, height=675)
        self.insertActivity('\n{}\n'.format(self.hello()))
        
        # プログレスバー
        self.progbarVal = tk.IntVar(value=0)
        self.progBar = ttk.Progressbar(self.root, orient='horizontal', variable=self.progbarVal, length=300, mode='determinate')
        self.progBar.place(x=600, y=675, width=400, height=25)

        self.enableWidgets('initial')

    def main(self):
        self.root.mainloop()        

if __name__ == "__main__":
    
    load_dotenv()
    if os.getenv('SENDGRID_API_KEY') is None:
        messagebox.showerror('環境変数エラー', 'APIキーが見つかりませんでした。\n.env ファイルを確認してください。')
        exit()

    sgoiMail().main()
