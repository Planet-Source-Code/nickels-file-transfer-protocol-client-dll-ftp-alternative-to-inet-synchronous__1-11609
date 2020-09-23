VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP Control"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock myTCPData 
      Left            =   120
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer myTimeout 
      Left            =   1095
      Top             =   150
   End
   Begin MSWinsockLib.Winsock myTCP 
      Left            =   615
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myObj As Connection
Public ftpEventFlag As Boolean
Public ftpEventError As Long
Public ftpEventReason As String

Public ftpDataConnected As Boolean
Public ftpDataSendComplete As Boolean
Public ftpDataServerClosed As Boolean
Public ftpWriteFileNum As Integer

Private inFileData() As Byte
Public ftpLogData As String

Public Sub ftpAddToLog(ByVal LogData As String)
    If Len(ftpLogData) + Len(LogData) > myObj.MaxLogSize Then
        ftpLogData = Mid(ftpLogData, Len(LogData))
    End If
    ftpLogData = ftpLogData & LogData
End Sub

Private Sub myTCP_Close()
    ftpEventReason = "Disconnected..."
    ftpEventError = False
End Sub

Private Sub myTCP_DataArrival(ByVal bytesTotal As Long)
    Dim inData As String
    Dim oneMsg As String
    myTCP.GetData inData, vbString, bytesTotal
    
    ftpAddToLog inData
    
    inData = Trim(TrimStrip(TrimStrip(Trim(inData), Chr(13)), Chr(10)))
    
    Do While Len(inData) > 0
        If InStr(inData, Chr(13)) > 0 Then
            oneMsg = TrimStrip(TrimStrip(Left(inData, InStr(inData, Chr(13)) - 1), Chr(13)), Chr(10))
            inData = Mid(inData, InStr(inData, Chr(13)) + 1)
        Else
            oneMsg = TrimStrip(TrimStrip(inData, Chr(13)), Chr(10))
            inData = ""
            End If
        Select Case Left(Trim(oneMsg), 3)
            Case "110"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "Restart marker reply."
                ftpEventFlag = False
            Case "125"
                'RAppeOk
                ftpEventError = ERR_OK
                ftpEventReason = "Data connection open; transfer starting."
                ftpEventFlag = False
            Case "150"
                'RAppeOk
                ftpEventError = ERR_OK
                ftpEventReason = Trim(oneMsg)
                ftpEventFlag = False
            Case "200"
                'RAlloOk
                ftpEventError = ERR_OK
                ftpEventReason = "The previous command was successfull."
                ftpEventFlag = False
            Case "202"
                'RAcctNotOk
                'RAlloNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The previous command used by this program is not implemented on the server and can not be used."
                ftpEventFlag = False
            Case "211"
                ftpEventError = ERR_OK
                ftpEventReason = "System status, or system help reply."
                ftpEventFlag = False
            Case "212"
                ftpEventError = ERR_OK
                ftpEventReason = "Directory status."
                ftpEventFlag = False
            Case "213"
                ftpEventError = ERR_OK
                ftpEventReason = "File status."
                ftpEventFlag = False
            Case "214"
                ftpEventError = ERR_OK
                ftpEventReason = "Help message."
                ftpEventFlag = False
            Case "215"
                ftpEventError = ERR_OK
                ftpEventReason = "System type."
                ftpEventFlag = False
            Case "220"
                ftpEventError = ERR_OK
                ftpEventReason = "System message."
                If Mid(Trim(oneMsg), 4, 1) <> "-" Then
                    ftpEventFlag = False
                    End If
            Case "225"
                'RAborOk
                ftpEventError = ERR_OK
                ftpEventReason = "The transfer is no longer in progress."
                ftpDataServerClosed = True
            Case "226"
                'RAborOk
                'RAppeOk
                ftpEventError = ERR_OK
                ftpEventReason = "The transfer is no longer in progress."
                ftpDataServerClosed = True
            Case "227"
                ftpEventError = ERR_OK
                ftpEventReason = Trim(oneMsg)
                ftpEventFlag = False
            Case "230"
                'RAcctOk
                'RUserOk
                ftpEventError = ERR_OK
                ftpEventReason = "The login is valid and you may proceed."
                If Mid(Trim(oneMsg), 4, 1) <> "-" Then
                    ftpEventFlag = False
                    End If
            Case "250"
                'RAppeOk
                'RCdupOk
                'RCwdOk
                'RDeleOk
                ftpEventError = ERR_OK
                ftpEventReason = "Requested file action okay, completed."
                If Mid(Trim(oneMsg), 4, 1) <> "-" Then
                    ftpEventFlag = False
                    End If
            Case "257"
                ftpEventError = ERR_OK
                ftpEventReason = Trim(oneMsg)
                ftpEventFlag = False
            Case "331"
                'RUserOk
                ftpEventError = ERR_OK
                ftpEventReason = "User name okay, need password."
                ftpEventFlag = False
            Case "332"
                'RUserNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "Need account for login."
                ftpEventFlag = False
            Case "350"
                ftpEventError = ERR_OK
                ftpEventReason = "Requested file action pending further information."
                ftpEventFlag = False
            Case "421"
                'RAborNotOk
                'RAlloNotOk
                'RAppeNotOk
                'RCdupNotOk
                'RCwdNotOk
                'RUserNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "Service not available, your connection has been terminated."
                ftpEventFlag = False
            Case "425"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "Can't open data connection with the server."
                'pTCPList(Index).Tag = ERR_Returned
                ftpEventFlag = False
            Case "426"
                'RAborNotOk
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "Connection closed; transfer aborted."
                ftpEventFlag = False
            Case "450"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The requested file/directory action was not taken, either the file/directory is unavailable or busy."
                ftpEventFlag = False
            Case "451"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The requested file/directory action was aborted due to an error in processing."
                ftpEventFlag = False
            Case "452"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The requested file/directory action was not completed due to insufficient space."
                ftpEventFlag = False
            Case "500"
                'RAborNotOk
                'RAcctNotOk
                'RAlloNotOk
                'RAppeNotOk
                'RCdupNotOk
                'RCwdNotOk
                'RUserNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The previous command used by this program is unrecognized by the server."
                ftpEventFlag = False
            Case "501"
                'RAbotNotOk
                'RAcctNotOk
                'RAlloNotOk
                'RAppeNotOk
                'RCdupNotOk
                'RCwdNotOk
                'RUserNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The previous command used by this program has different parameters or arguments then the servers command and can not be used."
                ftpEventFlag = False
            Case "502"
                'RAborNotOk
                'RAcctNotOk
                'RAppeNotOk
                'RCdupNotOk
                'RCwdNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The previous command used by this program is not implemented on the server and can not be used."
                ftpEventFlag = False
            Case "503"
                'RAcctNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The previous commands where not in correct order to the servers specifications and can not be used."
                ftpEventFlag = False
            Case "504"
                'RAlloNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The previous commands parameters where not implemented by the server and can not be used."
                ftpEventFlag = False
            Case "530"
                'RAcctNotOk
                'RAlloNotOk
                'RAppeNotOk
                'RCdupNotOk
                'RCwdNotOk
                'RUserNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "User not logged in correct."
                ftpEventFlag = False
            Case "532"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "You need an account for storing files."
                ftpEventFlag = False
            Case "550"
                'RAppeNotOk
                'RCdupNotOk
                'RCwdNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The requested file/directory action was aborted because the file/directory was not found or you don't have access."
                ftpEventFlag = False
            Case "551"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The requested file/directory action was aborted because the page type was unknown."
                ftpEventFlag = False
            Case "552"
                'RAppeNotOk
                ftpEventError = ERR_Returned
                ftpEventReason = "The requested file/directory action was aborted because it exceeded storage allocation."
                ftpEventFlag = False
            Case "553"
                'RAppeNotOk
                
                'was
                'ftpEventError = ERR_Returned
                'now is
                ftpEventError = ERR_OK
                ftpEventReason = "The requested file/directory action was aborted because the file/directory name was not allowed."
                ftpEventFlag = False
        End Select
        
        inData = Trim(TrimStrip(TrimStrip(Trim(inData), Chr(13)), Chr(10)))
        Loop
End Sub




Private Sub myTCPData_Close()
    If myTCPData.State <> 0 Then myTCPData.Close
    ftpDataConnected = False
End Sub

Private Sub myTCPData_Connect()
    ftpDataConnected = True
End Sub

Private Sub myTCPData_ConnectionRequest(ByVal requestID As Long)
    If myObj.ConnectionMode = "PORT" Then
        If myTCPData.State <> 0 Then myTCPData.Close
        myTCPData.Accept requestID
        ftpDataConnected = True
    End If
End Sub

Private Sub myTCPData_DataArrival(ByVal bytesTotal As Long)
    
    ReDim inFileData(1 To bytesTotal) As Byte
    myTCPData.GetData inFileData(), , bytesTotal
    
    If ftpWriteFileNum > 0 Then
        Put #ftpWriteFileNum, , inFileData
    End If
    
End Sub


Private Sub myTCPData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ftpDataConnected = False
End Sub

Private Sub myTCPData_SendComplete()
    ftpDataSendComplete = True
End Sub

Private Sub myTimeout_Timer()
    ftpEventError = ERR_Returned
    ftpEventReason = "The servers response has timed out."
    ftpEventFlag = False
End Sub


'###########################################################################################
'###########################################################################################
'###########################################################################################
'Common functions
'###########################################################################################
'###########################################################################################
'###########################################################################################
Private Function TrimStrip(ByVal TheStr As String, ByVal thechar As String) As String
    Dim i As Long
    Do Until Left(TheStr, 1) <> thechar
        If Left$(TheStr, 1) = thechar Then TheStr = Mid$(TheStr, 2)
    Loop
    TheStr = StrReverse(TheStr)
    Do Until Left(TheStr, 1) <> thechar
        If Left$(TheStr, 1) = thechar Then TheStr = Mid$(TheStr, 2)
    Loop
    TrimStrip = StrReverse(TheStr)
End Function


