VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'************************************************************************
'FTPClient DLL v1.0 Copyright 2000 By NeoText
'
'
'Support:
'   support@neotextsoftware.com
'
'   http://www.neotextsoftware.com
'
'
'Terms of Agreement:
' By using this source code, you agree to the following terms...
'  1) You may use this source code in personal projects and may compile
'     it into an .exe/.dll/.ocx and distribute it in binary format
'     freely and with no charge.
'  2) You MAY NOT redistribute this source code (for example to a
'     web site) without written permission from the original author.
'     Failure to do so is a violation of copyright laws.
'  3) You may link to this code from another website, provided it
'     is not wrapped in a frame.
'  4) The author of this code may have retained certain additional
'     copyright rights.If so, this is indicated in the author's
'     description.
'************************************************************************

Private ftpMaxLogSize As Long
Private ftpTimeOut As Long
Private ftpConnectionMode As String



Private myForm As New frmFTP

Public Enum TransferModes
    ftp_BINARY = 0
    ftp_ASCII = 1
End Enum

Public Enum ErrorReturns
    ERR_OK = 0
    ERR_Returned = 1
End Enum

Public Property Get ConnectionMode() As String
    ConnectionMode = UCase(Trim(ftpConnectionMode))
End Property
Public Property Let ConnectionMode(ByVal NewVal As String)
    ftpConnectionMode = UCase(Trim(NewVal))
End Property

Public Property Get TimeOut() As Long
    TimeOut = ftpTimeOut
End Property
Public Property Let TimeOut(ByVal NewVal As Long)
    ftpTimeOut = NewVal
End Property

Public Property Get MaxLogSize() As Long
    MaxLogSize = ftpMaxLogSize
End Property
Public Property Let MaxLogSize(ByVal NewVal As Long)
    ftpMaxLogSize = NewVal
End Property
Public Property Get LogData() As String
    LogData = myForm.ftpLogData
End Property

Private Sub Class_Initialize()
    
    MaxLogSize = 40000
    TimeOut = 20000
    ConnectionMode = "PASV"
    
    Set myForm.myObj = Me
    
    myForm.myTimeout.Enabled = False
    myForm.ftpEventFlag = False
    myForm.ftpEventError = ERR_OK
    myForm.ftpEventReason = ""
    myForm.ftpDataConnected = False
    myForm.ftpDataSendComplete = False
    myForm.ftpDataServerClosed = False
    myForm.ftpWriteFileNum = 0
End Sub


Private Sub SendDataFTPCommand(ByVal ftpCommand As String)

    myForm.ftpDataConnected = False

    Select Case ftpConnectionMode
    Case "PORT"
        Dim LinkIP As String
        Dim LinkPort As Long
    
        If myForm.myTCPData.State <> 0 Then myForm.myTCPData.Close
        myForm.myTCPData.LocalPort = 0
        myForm.myTCPData.Listen
    
        LinkIP = myForm.myTCPData.LocalIP
        LinkPort = myForm.myTCPData.LocalPort
        SendFTPCommand "PORT " & Replace(LinkIP, ".", ",") & "," & LinkPort \ 256 & "," & LinkPort Mod 256
    
        SendFTPCommand ftpCommand
    
        WaitForDataOpen
        
    Case Else
        
        SendFTPCommand "PASV"
        
        Dim pServer As String
        Dim pPHigh As Long
        Dim pPLow As Long
        Dim pPort As Long
        pServer = StrReverse(myForm.ftpEventReason)
        pServer = Mid(pServer, InStr(pServer, ")") + 1)
        pServer = Left(pServer, InStr(pServer, "(") - 1)
        pPLow = CLng(StrReverse(Left(pServer, InStr(pServer, ",") - 1)))
        pServer = Mid(pServer, InStr(pServer, ",") + 1)
        pPHigh = CLng(StrReverse(Left(pServer, InStr(pServer, ",") - 1)))
        pServer = Replace(StrReverse(Mid(pServer, InStr(pServer, ",") + 1)), ",", ".")
        pPort = (Val(pPHigh) * 256) + Val(pPLow)
                    
        If myForm.myTCPData.State <> 0 Then myForm.myTCPData.Close
        myForm.myTCPData.LocalPort = 0
        myForm.myTCPData.RemotePort = pPort
        myForm.myTCPData.RemoteHost = pServer
        
        myForm.myTCPData.Connect

        WaitForDataOpen

        SendFTPCommand ftpCommand

    End Select
    
   
End Sub
Public Function Connect(ByVal eServer As String, ByVal ePort As String, ByVal eLogin As String, ByVal ePass As String)

    If myForm.myTCP.State <> 0 Then myForm.myTCP.Close
    myForm.myTCP.RemoteHost = eServer
    myForm.myTCP.LocalPort = 0
    myForm.myTCP.RemotePort = ePort
    myForm.myTCP.Connect
    
    WaitForTCP
    
    SendFTPCommand "USER " + eLogin
                
    SendFTPCommand "PASS " + ePass
    
End Function

Public Function Disconnect()

    myForm.myTCP.SendData "QUIT" & vbCrLf

    If myForm.myTCP.State <> 0 Then myForm.myTCP.Close
    
End Function

Public Function Rename(ByVal SourceFileName As String, ByVal DestFileName As String) As String

    SendFTPCommand "RNFR " & SourceFileName
    
    If myForm.ftpEventError = ERR_OK Then
    
        SendFTPCommand "RNTO " & DestFileName
    
    End If

End Function

Public Function ChangeDirectory(ByVal ToFolder As String)
    
    SendFTPCommand "CWD " & ToFolder

End Function

Public Function MakeDirectory(ByVal NewFolder As String)

    SendFTPCommand "MKD " & NewFolder

End Function

Public Function RemoveFile(ByVal FileName As String)

    SendFTPCommand "DELE " & FileName

End Function

Public Function RemoveDirectory(ByVal FolderName As String)

    SendFTPCommand "RMD " & FolderName

End Function

Public Function TransferType(ByVal TransType As TransferModes) As String

    If TransType = ftp_BINARY Then
    
        SendFTPCommand "TYPE I"
    
    ElseIf TransType = ftp_ASCII Then
        
        SendFTPCommand "TYPE A"
    
    End If

End Function

Public Function ListContents(ByVal LocalFileName As String) As String
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open LocalFileName For Output As #FileNum
    Close #FileNum
    Open LocalFileName For Binary Access Write As #FileNum
    myForm.ftpWriteFileNum = FileNum
    
    SendDataFTPCommand "LIST"

    WaitForDataClose
    WaitForDataServerClosed

    Close #FileNum
    myForm.ftpWriteFileNum = 0

End Function

Public Function GetFile(ByVal RemoteFileName As String, ByVal LocalFileName As String) As String
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open LocalFileName For Output As #FileNum
    Close #FileNum
    Open LocalFileName For Binary Access Write As #FileNum
    myForm.ftpWriteFileNum = FileNum
    
    SendDataFTPCommand "RETR " & RemoteFileName

    WaitForDataClose
    WaitForDataServerClosed

    Close #FileNum
    myForm.ftpWriteFileNum = 0

End Function

Public Function PutFile(ByVal RemoteFileName As String, ByVal LocalFileName As String) As String
    
    Const PacketSize = 3072
    
    Dim inFileData() As Byte
    
    Dim cnt As Long
    Dim bytesCount As Long
    Dim bytesLeft As Long

    SendDataFTPCommand "STOR " & RemoteFileName
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open LocalFileName For Binary Access Read As #FileNum
    myForm.ftpWriteFileNum = FileNum
    
        
    bytesCount = LOF(FileNum) / PacketSize
    bytesLeft = LOF(FileNum) Mod PacketSize
        
    ReDim inFileData(1 To PacketSize) As Byte
    
    If bytesCount > 0 Then
        For cnt = 1 To bytesCount
            Get #FileNum, , inFileData()
            myForm.myTCPData.SendData inFileData()
            WaitForDataSend
        Next
    End If
    
    If bytesLeft > 0 Then
        ReDim inFileData(1 To bytesLeft) As Byte
        Get #FileNum, , inFileData()
        myForm.myTCPData.SendData inFileData()
        WaitForDataSend
    End If
        
    myForm.myTCPData.Close
    
    WaitForDataServerClosed
    
    Close #FileNum
    myForm.ftpWriteFileNum = 0

End Function

Private Sub SendFTPCommand(ByVal ftpCommand As String)
    myForm.ftpEventFlag = True

    myForm.myTCP.SendData ftpCommand & vbCrLf
    
    myForm.ftpAddToLog ftpCommand & vbCrLf
    
    myForm.myTimeout.Interval = ftpTimeOut
    myForm.myTimeout.Enabled = True
    
    Do Until myForm.ftpEventFlag = False
        DoEvents
    Loop

    myForm.myTimeout.Enabled = False

    If myForm.ftpEventError = ERR_Returned Then Err.Raise 8, "FTP Command", myForm.ftpEventReason
End Sub
Private Sub WaitForTCP()
    myForm.ftpEventFlag = True
    
    myForm.myTimeout.Interval = ftpTimeOut
    myForm.myTimeout.Enabled = True
    
    Do Until myForm.ftpEventFlag = False
        DoEvents
    Loop

    myForm.myTimeout.Enabled = False

    If myForm.ftpEventError = ERR_Returned Then Err.Raise 0, "FTP Command", myForm.ftpEventReason
End Sub

Private Sub WaitForDataOpen()
    
    Do Until myForm.ftpDataConnected = True
        DoEvents
    Loop

    If myForm.ftpEventError = ERR_Returned Then Err.Raise 0, "FTP Command", myForm.ftpEventReason
End Sub

Private Sub WaitForDataClose()
    
    Do Until myForm.ftpDataConnected = False
        DoEvents
    Loop

    If myForm.ftpEventError = ERR_Returned Then Err.Raise 0, "FTP Command", myForm.ftpEventReason
End Sub

Private Sub WaitForDataSend()
    
    Do Until myForm.ftpDataSendComplete = True
        DoEvents
    Loop

    myForm.ftpDataSendComplete = False
    
    If myForm.ftpEventError = ERR_Returned Then Err.Raise 0, "FTP Command", myForm.ftpEventReason
End Sub

Private Sub WaitForDataServerClosed()
    
    Do Until myForm.ftpDataServerClosed = True
        DoEvents
    Loop

    myForm.ftpDataServerClosed = False
    
    If myForm.ftpEventError = ERR_Returned Then Err.Raise 0, "FTP Command", myForm.ftpEventReason
End Sub



