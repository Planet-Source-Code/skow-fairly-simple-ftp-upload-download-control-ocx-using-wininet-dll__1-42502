VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ftpControl Example."
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Quick Download"
      Height          =   420
      Left            =   4005
      TabIndex        =   2
      Top             =   1815
      Width           =   2235
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Quick Upload"
      Height          =   420
      Left            =   810
      TabIndex        =   1
      Top             =   1800
      Width           =   2235
   End
   Begin ftp_Control_Demo.ftpControl ftpControl1 
      Height          =   1410
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2487
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This is a simple control example.
'  Distribute/use as you wish.
'   And yes, I will be making a way more hardout version
'     However, to keep things simple I have uploaded this version
'       as the other version will be much more complex to use.
'
' :)


Private Sub InitFTPControl()
    ' Set values common to upload/download
    With Me.ftpControl1
        .ftp_is_ascii_mode = True                           ' True for Text files, False for bin files
        .ftp_local_filename = "C:\MyFile.txt"               ' local file to download to, or upload from
        .ftp_password = "my_password"                       ' password to login with
        .ftp_remote_address = "ftp.mywebsite.com"           ' ftp site to connect to
        .ftp_remote_filename = "/public_html/MyFile.txt"    ' remote file to download or upload
        .ftp_remote_port = 21                               ' default FTP port
        .ftp_username = "my_username"                       ' username, anonymous for downloads if required
    End With
    End Sub

Private Sub cmdDownload_Click()
    Dim boolResult As Boolean
    Dim strResult As String
        
    InitFTPControl
    
    boolResult = ftpControl1.ftp_Download_quick(strResult)
    
    MsgBox "Result: " & strResult, vbInformation, IIf(boolResult = True, "Success!", "Failed.")
    End Sub

Private Sub cmdUpload_Click()
    Dim boolResult As Boolean
    Dim strResult As String
        
    InitFTPControl
    
    boolResult = ftpControl1.ftp_Upload_quick(strResult)
    
    MsgBox "Result: " & strResult, vbInformation, IIf(boolResult = True, "Success!", "Failed.")
    End Sub
