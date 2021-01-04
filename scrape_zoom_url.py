# モジュールのインポート
import win32com.client
import datetime
import re
import sqlite3


# 集計区間内のメールを選ぶ関数
def choose_period(mail):

    # とりあえず30日分
    wanna_show_timedelta = 30
    now = datetime.datetime.today()
    try:   
        temp_t = str(mail.ReceivedTime).partition('+')[0].partition('.')[0]
        received_time = datetime.datetime.strptime(temp_t, "%Y-%m-%d %H:%M:%S")
    except:
        # ReceivedTimeの変数を持たないメールは False を返す
        return False

    # 30日以内に受信したメールは True を返す
    return now - received_time < datetime.timedelta(days=wanna_show_timedelta)


# 受信メールフォルダを選ぶ関数
def choose_inbox(folder):
    return str(folder) in ['Inbox', 'inbox', '受信トレイ']


# 本文に'zoom.us'のhttpアドレスを含むメールのみTrueを返す関数
def contains_zoom(mail):
    body = mail.body
    return 'zoom.us' in body and 'http' in body


# zoomのリンクを抽出する関数
def extract_url(body):

    # すべてのurlを抽出
    url_pattern = 'https?://[\w/:%#\$&\?\(\)~\.=\+\-]+'
    temp_url_list = re.findall(url_pattern, body)

    # そこからzoomアドレスのみ抽出
    temp_zoom_url_list = list(filter(lambda url: 'zoom.us' in url, temp_url_list)) 

    # 重複を削除
    zoom_url_list = list(dict.fromkeys(temp_zoom_url_list))

    # URLの数を保存し、"|"で接続
    url_numbers = len(zoom_url_list)
    urls = '|'.join(zoom_url_list)

    return urls, url_numbers


# テーブル作成
def create_table(cursor):
    # CREATE
    """
    テーブルの中身
        -   受信日
        -   件名
        -   本文
        -   送信者アドレス
        -   zoomアドレスのかたまり
        -   zoomアドレスの個数（重複なし）
    """
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sample0 (
            received_time TEXT,
            mail_subject TEXT,
            sender TEXT,
            sender_address TEXT,
            url_list TEXT,
            url_numbers INTEGER
        );
    """)


# テーブルへ入力
def input_table(mails, connection, cursor):

    # zoomアドレスを含むメール１通ごとの内容をテーブルに書き込む
    for mail in mails:

        # 受信日をdatetime型に変型
        temp_t = str(mail.ReceivedTime).partition('+')[0].partition('.')[0]
        received_time = datetime.datetime.strptime(temp_t, "%Y-%m-%d %H:%M:%S")

        # zoomのURLのかたまりと個数を取り出す
        url_list, url_numbers = extract_url(mail.body)

        # エラー処理
        try:            
            # INSERT
            cursor.execute(
                "INSERT INTO sample0 VALUES (?, ?, ?, ?, ?, ?)",
                    (
                    mail.subject,
                    str(mail.sender)[1:-1],     # 両側の<>を取る
                    mail.senderEmailAddress,
                    received_time,
                    url_list,
                    url_numbers
                    )
            )

        except sqlite3.Error as e:
            print("sqlite3.Error occurred: ", e.args[0])    

    # COMMIT
    connection.commit()


# いらないzoomアドレスを弾く関数
def delete_elem(connection, cursor):
    
    # エラー処理
    try:
        cursor.execute("""
            DELETE FROM sample0 where url_list in (
                "https://zoom.us/",
                "https://zoom.us/support/download",
                "https://zoom.us/test"
            );
        """)
    except sqlite3.Error as e:
        print("sqlite3.Error occurred: ", e.args[0])

    # COMMIT
    connection.commit()
        

# テーブルを表示する関数
def show_table(cursor):

    # テーブルからzoomURLの個数とかたまりを行ごとに表示
    for row in cursor.execute("select url_list, url_numbers from sample0"):
        print("{} zoom url(S): ".format(row[1]))
        print(row[0])
        print("")


# 一連の処理をmain()にまとめる（関数にまとめたほうが速い？）
def main():
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    
    # メールアドレスごとにアカウントを分ける
    accounts = outlook.Folders

    # データベース接続とカーソル作成
    connection = sqlite3.connect('sample0.db')
    cursor =  connection.cursor()

    # テーブルを作成
    create_table(cursor)

    # メールアドレスごとに処理
    for account in accounts:

        # 受信メールフォルダを選択
        folders = account.Folders
        # 受信フォルダはひとつなのでリストの一つ目を選択
        inbox = list(filter(choose_inbox, folders))[0]
        
        # zoomアドレスを持つメールを集計
        all_items = inbox.Items
        items = list(filter(choose_period, all_items))
        zoom_mails = list(filter(contains_zoom, items))

        # テーブルに入力
        input_table(zoom_mails, connection, cursor)


    # 特定のzoomアドレスの要素は削除
    delete_elem(connection, cursor)

    # 出力
    show_table(cursor)

    # Close
    connection.close()


main()