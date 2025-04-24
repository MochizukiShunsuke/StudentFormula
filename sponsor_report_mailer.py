import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import logging
from email.header import Header
import os  # ファイル名を取得するために使用

# 入力：年度と月、PDFパス
year = input("年度を入力してください: ")
month = input("月を入力してください: ")
filepath = input("添付するPDFのフルパスを入力してください: ")

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# SMTPサーバーの設定
smtp_server = 'mail.kanazawa-formula.com'
port = 465

# CSVファイルからデータを読み込む
data = pd.read_csv('sponsor_data.csv')

# 各行をループして送信者情報を取得し、メールを送信
for index, row in data.iterrows():
    sender_email = row['送信元アドレス']
    password = row['送信元アドレスパスワード']
    receiver_email = row['送信先アドレス']

    try:
        # SMTPサーバーに接続
        server = smtplib.SMTP_SSL(smtp_server, port)
        server.login(sender_email, password)

        # メールテンプレートの作成
        template = """
{company_name1}
{company_name2} 様

お世話になっております。
{team_name}の{your_name1}({your_name2})と申します．
{year}年度チームでは，{section_name}を担当しております．

{month}月のフォーミュラ研究会の近況活動報告をさせていただきます．
以下の1点を添付させていただきましたので，よろしければご覧ください．
・{year}年度{month}月近況活動報告書

今後とも，金沢大学フォーミュラ研究会を何卒よろしくお願いいたします． 


-- 
{affiliation}
{your_name1}({your_name2})
{address}
{team_name}
携帯番号(個人)：{TEL}
E-mail: {sender_email} （個人）
E-mail: office@kanazawa-formula.com （チーム）
URL: http://kanazawa-formula.com
        """

        # メールの内容を作成
        subject = "【金沢大学フォーミュラ研究会】近況活動報告のお知らせ"
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject

        # メール本文をUTF-8で設定
        body = template.format(
            sender_email=row['送信元アドレス'],
            company_name1=row['会社名'],
            company_name2=row['送信先名前'],
            your_name1=row['送信元名前'],
            your_name2=row['なまえ'],
            section_name=row['セクション'],
            affiliation=row['所属'],
            address=row['住所'],
            team_name=row['チーム名'],
            TEL = row['携帯番号(個人)'],
            year=year,
            month=month
        )
        message.attach(MIMEText(body, 'plain', 'utf-8'))

        # PDFファイルを添付
        filename = os.path.basename(filepath)  # ファイル名だけを抽出
        try:
            with open(filepath, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                # ファイル名をUTF-8でエンコードしてContent-Dispositionを設定
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{Header(filename, "utf-8").encode()}"'
                )
                message.attach(part)
        except FileNotFoundError:
            logging.error(f"添付ファイルが見つかりません: {filepath}")
            continue  # 添付ファイルがない場合、このメールの送信をスキップ

        # メールを送信
        server.sendmail(sender_email, receiver_email, message.as_string())
        logging.info(f"メールを送信しました: {receiver_email}")

    except Exception as e:
        logging.error(f"{receiver_email}へのメール送信に失敗しました: {e}")

    finally:
        # サーバーの接続を必ず閉じる
        server.quit()