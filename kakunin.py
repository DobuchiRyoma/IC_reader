import PySimpleGUI as sg
import io
import msoffcrypto
import openpyxl as excel
import os
import nfc
import os
import threading
import csv
import nfc.clf.rcs380

PASORI_S380_PATH = 'usb'   # PaSoRi RC-S380
outfile = 'temporary.xlsx' # 一時的な出力ファイル名
nonyuList = []
allList = []
nonyuLine = 0
gakubuLine = 0
isReading =False

#デザインテーマの設定
sg.theme("DarkGray13")

#ウィンドウの部品とレイアウト
layout = [
    [sg.Text('操作手順')],   
    [sg.Text('１：納入状況を示したエクセルファイルを選びパスワードを入力')],   
    [sg.Text('２：準備完了ボタンを押す')],   
    [sg.Text('３：読み取り開始ボタンを押す')],  
    [sg.Text('※：読み取り中にエラーが出たときは読み取り開始ボタンをもう一度押すか再起動')],  
    [sg.Text('イベント名'),sg.In(key='-IN-')],
    [sg.Text('ファイル', size=(10, 1)), sg.Input(), sg.FileBrowse('ファイルを選択', key='inputFilePath')],
    [sg.Text('パスワード', size=(10, 1)), sg.InputText('', size=(10, 1), key='passward')],
    [sg.Button('準備完了', key='ready')],[sg.Button('読み取り開始', key='read'),sg.Text('学生証を忘れた人の検索用'),sg.In(key='num', size=(10, 1)),sg.Button('検索', key='search')],
    [sg.Output(size=(80,20))]
]

#ウィンドウの生成
window = sg.Window('学友会費納入確認', layout)

#読み込んだ結果を返す
def sc_from_raw(sc):
    return nfc.tag.tt3.ServiceCode(sc >> 6, sc & 0x3f)

#マルチスレッドによる学生証タッチの受付
def on_connect(tag):
    with open(values['-IN-']+".csv", 'a', newline='') as f:
        writer = csv.writer(f)
        sc1 = sc_from_raw(0x200B)
        bc1 = nfc.tag.tt3.BlockCode(0, service=0)
        bc2 = nfc.tag.tt3.BlockCode(1, service=0)
        block_data = tag.read_without_encryption([sc1], [bc1, bc2])
        index = 0
        for i,studentNonyuInfo in enumerate(nonyuList):
            if studentNonyuInfo[0][1:] == block_data[17:23].decode("utf-8"):
                index = i
                print(studentNonyuInfo)              
        writer.writerow(allList[index])
        f.close()
    return True

def func():
    
        while True:
            with nfc.ContactlessFrontend(PASORI_S380_PATH) as clf:
                clf.connect(rdwr={
                    'on-connect': on_connect,
                })
                

#イベントループ
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED: #ウィンドウのXボタンを押したときの処理
        break

    if event == 'ready': #「読み取り」ボタンが押されたときの処理
        #パスワード付きExcelファイルを読み込む
        if values['-IN-'] == "" or values['-IN-'] == None:
            print("イベント名を入力してください")
        else:
            decrypted = io.BytesIO()
            #誤ったファイルを指定された場合
            if not (values['inputFilePath'].split(".")[-1] == "csv" or values['inputFilePath'].split(".")[-1] == "xlsx"):
                print("csvファイルかxlsxファイルを指定してください")
            else:
                try:
                    with open(values['inputFilePath'], 'rb') as fp:
                        msfile = msoffcrypto.OfficeFile(fp)
                        msfile.load_key(password=values["passward"])
                        msfile.decrypt(decrypted)
                        with open(outfile, 'wb') as fp:
                            fp.write(decrypted.getbuffer())
                except:
                    print("おそらくパスワードが違います")
                is_file = os.path.isfile(outfile)
                if is_file:
                    wbook = excel.load_workbook(filename=outfile)
                    sheets = wbook.sheetnames
                    for sheetName in sheets:
                        sheet = wbook[sheetName]
                        for i,cell in enumerate(sheet[1]):
                            if "納入" in str(cell.value):#納入を含む列なら
                                nonyuLine = i
                            if "学部名" in str(cell.value):
                                gakubuLine = i
                        for row in sheet.values:
                            if row[0] is None:continue
                            nonyuList.append([row[0],row[gakubuLine],row[nonyuLine]])
                            allList.append(row)
                    os.remove(outfile)    
                    isReading = True
                    print("準備完了！！")
                else:
                    print("エクセルファイルの解読に失敗しました")



    if event == "read":
        if isReading:
            window.start_thread(lambda: func(), ('-THREAD-', '-THEAD ENDED-'))
        else:
            print("準備を完了して下さい")
    
    if event == "search":
        if values['num'] == "" or values['num'] == None:
            print("学籍番号を入力して下さい")
        else:
            with open(values['-IN-']+".csv", 'a', newline='') as f:
                writer = csv.writer(f)
                index = 0
                for i,studentNonyuInfo in enumerate(nonyuList):
                    if studentNonyuInfo[0][1:] == values['num']:
                        index = i
                        print(studentNonyuInfo)              
                writer.writerow(allList[index])
        f.close()

window.close()

#コンパイルの際のコマンド
#pyinstaller --noconsole --icon=tut.ico kakunin.py