import datetime
import asyncio
import pip
import discord
import gspread
import json

# TODO メソッド化がよくわからん。asyncを意識しないとダメっぽい
# エラー解決するのに時間かけたくないのでとりあえずズラズラ書く

client = discord.Client()

TOKEN = 'NTUwNDk0Njc3NTU2Nzg5MjY4.D1jPrg.cnsFgano2XEotuHKF7dhu8FC1xs'
GOOGLE_TOKEN = 'AIzaSyC5CxECQlr860m4auZ7OO9p4DbmbwanZKk'
GOOGLE_JSON_FILE = 'MyProject-e027354dd4ff.json'
ENABLE_CHANNEL_ID = '550493878843867169' # 本番に合わせて変更
SHEET_ID_MISOJI = '1Kf8tZmfp58DCdVgsvz0H4TS2ED9huvcQJe3zTMh7fCQ'
SHEET_ID = '1cUPftGJMU21A8Y1UpTSCKyfz99MIJ0yP-kjfsONR77k'
SHEET_WORK_BOOK_NAME_ANSWER = 'フォームの回答 1'
SHEET_WORK_BOOK_NAME_MEIBO = '名簿'
NEW_LINE = '\n'

_builtMsg = None
_gc = None

#####################以下discord#####################s
# 起動時に通知してくれる処理
@client.event
async def on_ready():
    _gc = connectSpreadSheets()
    print('ログインしました')

# messsage受信時の処理
@client.event
async def on_message(message):
    if message.channel.id != ENABLE_CHANNEL_ID: # Channel idは本番環境に合わせて変更
        return
    if message.content == 'まー':
        await client.send_message(message.channel, '\U0001F4A9')
        return
    if message.content.startswith('bot'):
        if message.content.count('ヘルプ'):
            msg = '我輩はみそじちゃんである！キーワード一覧を教えてやるのである！'
            msg += NEW_LINE
            msg += 'bot 点呼：点呼用のURLを投下'
            msg += NEW_LINE
            msg += 'bot 確認：自分が入力したかを確認する'
            msg += NEW_LINE
            msg += 'bot 完了：点呼が完了している人一覧を表示させる'
            msg += NEW_LINE
            msg += 'bot 呼出：点呼が完了していない人にメンション飛ばす'
            await client.send_message(message.channel, msg)
            return
        if message.content.count('点呼'):#fix
            if message.content.count('：'):
                return
            await client.send_message(message.channel, '我輩はみそじちゃんである！')
            await client.send_message(message.channel, 'https://goo.gl/forms/X6Zga53tKYh9D6s42')
            await client.send_message(message.channel, '木曜日１８時までにお願いするのである！')
            return
        if message.content.count('確認'):#fix
            if message.content.count('：'):
                return
            await client.send_message(message.channel, '我輩はみそじちゃんである！')
            if isComplete(connectSpreadSheets(), message.author.id):
                msg = message.author.mention + " 入力済みなのである！"
            else:
                msg = message.author.mention + " まだなのである！早めの入力をお願いするのである！"
            await client.send_message(message.channel, msg)
            return
        if message.content.count('完了'):#fix
            if message.content.count('：'):
                return
            await client.send_message(message.channel, '我輩はみそじちゃんである！\n入力済みユーザを称えるのである！')
            userList = getAnswerUserList(connectSpreadSheets())
            _builtMsg = ''
            for name in userList:
                _builtMsg += name + NEW_LINE
            await client.send_message(message.channel, _builtMsg + 'よくやったアル！')
            return
        if message.content.count('呼出'):#fix
            if message.content.count('：'):
                return
            await client.send_message(message.channel, '我輩はみそじちゃんである！' + '未入力者を貼り出すのである！')
            userList = getAnswerUserList(connectSpreadSheets())
            userDict = getRosterData(connectSpreadSheets())
            _builtMsg = ''
            for k, v in userDict.items():
                if v not in userList:
                    _builtMsg += '<@!'+ k + '> ' + v
                    _builtMsg += NEW_LINE

            _builtMsg += '木曜日１８時までに入力をよろしくなのであーる！'
            await client.send_message(message.channel, _builtMsg)
            await client.send_message(message.channel, 'https://goo.gl/forms/X6Zga53tKYh9D6s42')
            return
    else:
        return

def connectSpreadSheets():
    global _gc
    if _gc == None:
        #ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
        from oauth2client.service_account import ServiceAccountCredentials 
        #2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        #認証情報設定
        #ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
        credentials = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_JSON_FILE, scope)
        #OAuth2の資格情報を使用してGoogle APIにログインします。
        return gspread.authorize(credentials)
    else:
        if isConnected(_gc):
            return _gc
        else:
            _gc = None
            return connectSpreadSheets()

#未入力検知メモ
#名簿のE列(ID)と一致するA列(名前)を取得
#フォームの回答 1のB列と一致するかを調べる
def getRosterData(gc):
    #共有設定したスプレッドシートのシート1を開く
    worksheet = gc.open_by_key(SHEET_ID).worksheet(SHEET_WORK_BOOK_NAME_MEIBO)
    nameList = worksheet.col_values(1)
    idList = worksheet.col_values(5)
    nameList.pop(0)
    idList.pop(0)
    return dict(zip(idList, nameList))

def getAnswerUserList(gc):
    worksheet = gc.open_by_key(SHEET_ID).worksheet(SHEET_WORK_BOOK_NAME_ANSWER)
    colList = worksheet.col_values(2) #B列を全取得
    colList.pop(0)#ヘッダテキストを削除
    return colList

def isComplete(gc, userId):
    # idに紐づく名前取得
    worksheet = gc.open_by_key(SHEET_ID).worksheet(SHEET_WORK_BOOK_NAME_MEIBO)
    cell = worksheet.find(userId)
    name = worksheet.cell(cell.row, 1).value
    # 名前が回答シートにあるか
    answerList = getAnswerUserList(gc)
    return name in answerList

def isConnected(gc):
    try:
        gc.open_by_key(SHEET_ID)
        return True
    except Exception as e:
        print(e)
        return False
    return    

# BOTを実行
client.run(TOKEN)