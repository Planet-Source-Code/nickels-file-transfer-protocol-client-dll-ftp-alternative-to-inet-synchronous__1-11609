Attribute VB_Name = "modMain"
Option Explicit

Sub Example()
Dim myFTP As New FTPConnection


myFTP.ConnectionMode = "PORT"

myFTP.Connect "24.26.169.236", 21, "chrome", "calsrfF7"

myFTP.ChangeDirectory "/d/inetpub"

myFTP.TransferType ftp_BINARY

'myFTP.PutFile "/d/InetPub/testingFTP.mdb", App.Path & "\ThaBBoards.mdb"

'myFTP.TransferType ftp_ASCII

'myFTP.ListContents App.Path & "\list.txt"

'myFTP.TransferType ftp_BINARY

'myFTP.GetFile "/d/InetPub/testingFTP.mdb", App.Path & "\testingThaFTP.mdb"

'myFTP.MakeDirectory "A NewFoLdEr"

myFTP.RemoveFile "testingFTP.mdb"

myFTP.RemoveFolder "A NewFoLdEr"

myFTP.ListContents App.Path & "\list.txt"


MsgBox myFTP.LogData

End Sub
