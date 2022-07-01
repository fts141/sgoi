# SGOI - SendGrid Operation Interface

import os
import logging
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
from dotenv import load_dotenv


class sgoiMail():

    def __init__(self):
        self.from_email = os.getenv('SENDGRID_MAIL_FROM')
        self.to_emails = os.getenv('SENDGRID_MAIL_TO')
        self.subject = 'TEST MAIL'
        self.html_content = 'これはテストメールです。<br />削除願います。'

    def send(self):
        try:
            sg = SendGridAPIClient(os.getenv('SENDGRID_API_KEY'))
            message = {'from_email': self.from_email, 'to_emails': self.to_emails, 'subject': self.subject, 'html_content': self.html_content}
            response = sg.send(Mail(**message))
            logging.info('メール送信：{} -> {}'.format(self.to_emails, response.status_code))
            logging.debug('== headers ==\n{}\n== body ==\n{}\n'.format(response.headers, response.body))
        except Exception as e:
            logging.warning('メール送信に失敗：{}'.format(e))


if __name__ == "__main__":
    
    load_dotenv()
    sm = sgoiMail()
    sm.send()
