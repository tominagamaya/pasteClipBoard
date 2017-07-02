import win32com.client
from datetime import datetime,timezone

# 対象の日付を設定
startDate = datetime(2017,7,1,0,0,0,0,timezone.utc)
endDate = datetime(2017,7,3,0,0,0,0,timezone.utc)

object = win32com.client.Dispatch("Outlook.Application")
mapi = object.GetNamespace("MAPI")
inbox = mapi.GetDefaultFolder(win32com.client.constants.olFolderInbox)

# 検索対象フォルダ
folder = inbox.Folders['レポート']
for i in folder.Items:
     if (endDate >= i.SentOn) and (i.SentOn >= startDate):
         if(i.Body.find('エラー') != -1):
             print(i.SentOn)
             print(i.Subject + "\n")
             # フラグを設定
             i.FlagIcon = 6
             i.Save()
