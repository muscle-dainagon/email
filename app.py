"""メールを自動で送信する"""

from email.mime.text import MIMEText
import smtplib
import json


class EmailSender:
    """メールを送信するクラス

    Attributes:
        from_email: Gamilのアドレス（メールの送信元）
        password: Gmailのアプリパスワード
        to_email: メールの宛先
        charset: 文字コード
        subject: 件名
        message: メールの本文

    """
    def __init__(self, to_email: str, subject: str, message: str):
        ## jsonから必要な情報を取得する
        settings_file = open("settings.json", "r")
        settings_data = json.load(settings_file)

        self.from_email = settings_data["from_email"]
        self.password = settings_data["password"]
        self.to_email = to_email
        self.charset = "ISO-2022-JP" # 日本語の文字エンコーディングの一つで、特に電子メールで使用される
        self.subject = subject
        self.message = message

    def send(self):
        """送信するメールの設定"""
        msg = MIMEText(self.message.encode(self.charset), "plain", self.charset)
        msg["Subject"] = self.subject
        msg["From"] = self.from_email
        msg["To"] = self.to_email

        ## Gmailのsmtp経由で送信
        smtp = smtplib.SMTP("smtp.gmail.com", 587)
        smtp.ehlo()
        smtp.starttls()
        smtp.login(self.from_email, self.password)
        smtp.send_message(msg)
        smtp.close()
        print("送信完了")


def main():
    settings_file = open("settings.json", "r")
    settings_data = json.load(settings_file)

    to_email = settings_data["test_email"]
    subject = "Pythonによるテストメール"
    message = "テストメッセージ"
    mailer = EmailSender(to_email, subject, message)
    mailer.send()

if __name__ == "__main__":
    main()