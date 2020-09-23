VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim ftpObj

Set ftpObj = CreateObject("FTPClient.Connection")

'connect to ftp server
ftpObj.Connect "127.0.0.1", 21, "anonymous", "testing"

'set connection mode to passive
ftpObj.connectionmode = "PASV"

'set the transfertype to ASCII
ftpObj.transfertype ftp_ASCII

'list the current directory to C:\list.dat
ftpObj.listcontents "C:\list.dat"

'disconnect from the ftp server
ftpObj.disconnect

'deinitialize object
Set ftpObj = Nothing

End Sub
