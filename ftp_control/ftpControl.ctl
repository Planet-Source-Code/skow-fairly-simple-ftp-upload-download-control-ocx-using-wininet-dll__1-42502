VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ftpControl 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1410
   ScaleWidth      =   6870
   ToolboxBitmap   =   "ftpControl.ctx":0000
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel Transfer"
      Height          =   375
      Left            =   2505
      TabIndex        =   2
      Top             =   900
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pbTransferStatus 
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "ftpControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  Before you go saying it has too much in here, listen to what I gots to say
'
'   This is framework for a much more complex control
'   The final control will have full progress bar, Cancel button
'   and events to pass back current %, file pos, file len etc
'
'   Won't be uploading it though, if you want that do it yourself :)
'
'

Private Type ftpConnectionType
    strAddress As String
    intPort As Integer
    strLocalFilename As String
    strRemoteFilename As String
    strUsername As String
    strPassword As String
    boolAsciiMode As Boolean
    End Type
    
Private Type ftpCurrentXfer
    strAddress As String
    strRemoteFilename As String
    strLocalFilename As String
    strUsername As String
    strPassword As String
    boolAsciiMode As Boolean
    intPort As Integer
    dblFileSize As Double
    dblFileLoc As Double
    intFileNo As Integer
    End Type
    
Private ftpControlVals As ftpConnectionType

Private session As Long
Private server As Long
Private Transfer As Long
Private Adresa As String
Private ID As String
Private Pass As String
Private Port As Integer
Private strPath As String
Private adr As String
Private Klic As String
 
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
 
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal pub_lngInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Sub InternetSetStatusCallback Lib "wininet.dll" (ByVal pub_lngInternetSession As Long, ByVal lpfnInternetCallback As Long)
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal pub_lngInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetQueryDataAvailable Lib "wininet.dll" (ByVal hInet As Long, dwAvail As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
 
Private Declare Function InternetTimeToSystemTime Lib "wininet.dll" (ByVal lpszTime As String, ByRef pst As SYSTEMTIME, ByVal dwReserved As Long) As Long
         
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Long
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String, ByVal fdwAccess As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Long
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
                  
' Use registry access settings.
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
 
' Number of the TCP/IP port on the server to connect to.
Private Const INTERNET_INVALID_PORT_NUMBER = 0
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_DEFAULT_GOPHER_PORT = 70
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443
Private Const INTERNET_DEFAULT_SOCKS_PORT = 1080
 
' Type of service to access.
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3
 
' Brings the data across the wire even if it locally cached.
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const ERROR_NO_MORE_FILES = 18
 
Private Const FTP_TRANSFER_TYPE_UNKNOWN As Long = &H0 '0x00000000
Private Const FTP_TRANSFER_TYPE_ASCII As Long = &H1 '0x00000001
Private Const FTP_TRANSFER_TYPE_BINARY  As Long = &H2 '0x00000002
 
' The possible values for the lInfoLevel parameter include:
Private Const HTTP_QUERY_CONTENT_TYPE = 1
Private Const HTTP_QUERY_CONTENT_LENGTH = 5
Private Const HTTP_QUERY_EXPIRES = 10
Private Const HTTP_QUERY_LAST_MODIFIED = 11
Private Const HTTP_QUERY_PRAGMA = 17
Private Const HTTP_QUERY_VERSION = 18
Private Const HTTP_QUERY_STATUS_CODE = 19
Private Const HTTP_QUERY_STATUS_TEXT = 20
Private Const HTTP_QUERY_RAW_HEADERS = 21
Private Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Private Const HTTP_QUERY_FORWARDED = 30
Private Const HTTP_QUERY_SERVER = 37
Private Const HTTP_QUERY_USER_AGENT = 39
Private Const HTTP_QUERY_SET_COOKIE = 43
Private Const HTTP_QUERY_REQUEST_METHOD = 45
 
' Add this flag to the about flags to get request header.
Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
 
' flags for InternetOpenUrl
Private Const INTERNET_FLAG_RAW_DATA = &H40000000
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Private Const INTERNET_FLAG_TRANSFER_ASCII = &H1&
Private Const INTERNET_FLAG_TRANSFER_BINARY = &H2&
 
' flags for InternetOpen
Private Const INTERNET_FLAG_ASYNC = &H10000000
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_DONT_CACHE = &H4000000
Private Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000
Private Const INTERNET_FLAG_OFFLINE = &H1000000
 
Private Type INTERNET_ASYNC_RESULT
    dwResult As Long
    dwError As Long
    End Type
 
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000     ' don't write this item to the cache
Private Const INTERNET_STATUS_RESOLVING_NAME = 10
Private Const INTERNET_STATUS_NAME_RESOLVED = 11
Private Const INTERNET_STATUS_CONNECTING_TO_SERVER = 20
Private Const INTERNET_STATUS_CONNECTED_TO_SERVER = 21
Private Const INTERNET_STATUS_SENDING_REQUEST = 30
Private Const INTERNET_STATUS_REQUEST_SENT = 31
Private Const INTERNET_STATUS_RECEIVING_RESPONSE = 40
Private Const INTERNET_STATUS_RESPONSE_RECEIVED = 41
Private Const INTERNET_STATUS_CTL_RESPONSE_RECEIVED = 42
Private Const INTERNET_STATUS_PREFETCH = 43
Private Const INTERNET_STATUS_CLOSING_CONNECTION = 50
Private Const INTERNET_STATUS_CONNECTION_CLOSED = 51
Private Const INTERNET_STATUS_HANDLE_CREATED = 60
Private Const INTERNET_STATUS_HANDLE_CLOSING = 70
Private Const INTERNET_STATUS_REQUEST_COMPLETE = 100
Private Const INTERNET_STATUS_REDIRECT = 110
Private Const INTERNET_STATUS_STATE_CHANGE = 200
Private Const INTERNET_ERROR_BASE = 12000
 
Private Const ERROR_INTERNET_OUT_OF_HANDLES = 12001
Private Const ERROR_INTERNET_TIMEOUT = 12002
Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Private Const ERROR_INTERNET_INTERNAL_ERROR = 12004
Private Const ERROR_INTERNET_INVALID_URL = 12005
Private Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = 12006
Private Const ERROR_INTERNET_NAME_NOT_RESOLVED = 12007
Private Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = 12008
Private Const ERROR_INTERNET_INVALID_OPTION = 12009
Private Const ERROR_INTERNET_BAD_OPTION_LENGTH = 12010
Private Const ERROR_INTERNET_OPTION_NOT_SETTABLE = 12011
Private Const ERROR_INTERNET_SHUTDOWN = 12012
Private Const ERROR_INTERNET_INCORRECT_USER_NAME = 12013
Private Const ERROR_INTERNET_INCORRECT_PASSWORD = 12014
Private Const ERROR_INTERNET_LOGIN_FAILURE = 12015
Private Const ERROR_INTERNET_INVALID_OPERATION = 12016
Private Const ERROR_INTERNET_OPERATION_CANCELLED = 12017
Private Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = 12018
Private Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = 12019
Private Const ERROR_INTERNET_NOT_PROXY_REQUEST = 12020
Private Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = 12021
Private Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = 12022
Private Const ERROR_INTERNET_NO_DIRECT_ACCESS = 12023
Private Const ERROR_INTERNET_NO_CONTEXT = 12024
Private Const ERROR_INTERNET_NO_CALLBACK = 12025
Private Const ERROR_INTERNET_REQUEST_PENDING = 12026
Private Const ERROR_INTERNET_INCORRECT_FORMAT = 12027
Private Const ERROR_INTERNET_ITEM_NOT_FOUND = 12028
Private Const ERROR_INTERNET_CANNOT_CONNECT = 12029
Private Const ERROR_INTERNET_CONNECTION_ABORTED = 12030
Private Const ERROR_INTERNET_CONNECTION_RESET = 12031
Private Const ERROR_INTERNET_FORCE_RETRY = 12032
Private Const ERROR_INTERNET_INVALID_PROXY_REQUEST = 12033
Private Const ERROR_INTERNET_NEED_UI = 12034
 
Private Const ERROR_INTERNET_HANDLE_EXISTS = 12036
Private Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = 12037
Private Const ERROR_INTERNET_SEC_CERT_CN_INVALID = 12038
Private Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = 12039
Private Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = 12040
Private Const ERROR_INTERNET_MIXED_SECURITY = 12041
Private Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = 12042
Private Const ERROR_INTERNET_POST_IS_NON_SECURE = 12043
Private Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = 12044
Private Const ERROR_INTERNET_INVALID_CA = 12045
Private Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = 12046
Private Const ERROR_INTERNET_ASYNC_THREAD_FAILED = 12047
Private Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = 12048
 
'//
'// FTP API errors
'//
 
Private Const ERROR_FTP_TRANSFER_IN_PROGRESS = 12110
Private Const ERROR_FTP_DROPPED = 12111
 
'//
'// gopher API errors
'//
 
Private Const ERROR_GOPHER_PROTOCOL_ERROR = 12130
Private Const ERROR_GOPHER_NOT_FILE = 12131
Private Const ERROR_GOPHER_DATA_ERROR = 12132
Private Const ERROR_GOPHER_END_OF_DATA = 12133
Private Const ERROR_GOPHER_INVALID_LOCATOR = 12134
Private Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = 12135
Private Const ERROR_GOPHER_NOT_GOPHER_PLUS = 12136
Private Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = 12137
Private Const ERROR_GOPHER_UNKNOWN_LOCATOR = 12138
 
'//
'// HTTP API errors
'//
 
Private Const ERROR_HTTP_HEADER_NOT_FOUND = 12150
Private Const ERROR_HTTP_DOWNLEVEL_SERVER = 12151
Private Const ERROR_HTTP_INVALID_SERVER_RESPONSE = 12152
Private Const ERROR_HTTP_INVALID_HEADER = 12153
Private Const ERROR_HTTP_INVALID_QUERY_REQUEST = 12154
Private Const ERROR_HTTP_HEADER_ALREADY_EXISTS = 12155
Private Const ERROR_HTTP_REDIRECT_FAILED = 12156
Private Const ERROR_HTTP_NOT_REDIRECTED = 12160               '// BUGBUG
 
Private Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = 12157   '// BUGBUG
Private Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = 12158    ' // BUGBUG
Private Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = 12159      '// BUGBUG
 
Private Const INTERNET_ERROR_LAST = 12159
 
Private Const MAX_PATH = 260
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2
 
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_ALWAYS = 4
 
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
        End Type
 
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
        End Type
        
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
        End Type
        
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
        End Type
        
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
        End Type
 
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpfilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpfilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FileTimeToSystemTime& Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME)
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
 

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const LANG_USER_DEFAULT = &H400&
 
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 
Private Const RAS_MAXENTRYNAME As Integer = 256
Private Const RAS_MAXDEVICETYPE As Integer = 16
Private Const RAS_MAXDEVICENAME As Integer = 128
Private Const RAS_RASCONNSIZE As Integer = 412

Private Type RASCONN
        dwSize As Long
        hRasConn As Long
        szEntryName(RAS_MAXENTRYNAME) As Byte
        szDeviceType(RAS_MAXDEVICETYPE) As Byte
        szDeviceName(RAS_MAXDEVICENAME) As Byte
        End Type
 
Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (udtRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

' Events to be implimented in my own version.. heh, do the rest yourself.
'
'Public Event XferStatus(intPercentComplete As Integer, lngBytesSent As Long, lngBytesRecieved As Long, lngBytesTotal As Long)
'Public Event XferStart(lngBytesTotal As Long)
'Public Event XferComplete(lngBytesTotal As Long)
'Public Event XferFailed(lngBytesSent As Long, lngBytesTotal As Long)
'Public Event XferCancelled(lngBytesSent As Long, lngBytesTotal As Long)


Private Function UpdateStatus(lngFilePos, lngFileLen, strText)
    On Error Resume Next
    Dim IntPos As Integer
    If lngFileLen <> 0 Then
        ' Update the status
        IntPos = Int((lngFilePos / lngFileLen) * 100)
        pbTransferStatus.Max = 100
        pbTransferStatus.Min = 0
        pbTransferStatus.Value = IntPos
        lblStatus.Caption = strText & " (" & IntPos & "%)"
    Else
        pbTransferStatus.Max = 100
        pbTransferStatus.Min = 0
        pbTransferStatus.Value = 0
        lblStatus.Caption = strText
    End If
    DoEvents
    End Function
    

Public Property Get ftp_remote_port() As Integer
Attribute ftp_remote_port.VB_Description = "sets remote port (default 21) for xfers"
    ftp_remote_port = ftpControlVals.intPort
    End Property

Public Property Let ftp_remote_port(ByVal vNewValue As Integer)
    ftpControlVals.intPort = vNewValue
    PropertyChanged "ftp_remote_port"
    End Property

Private Sub cmdCancel_Click()
    ' Not implimented for this control version.
    ' Will abort the FtpOpenFile() method.
    ' Not useable for FtpGetFile() / FtpWriteFile()  methods
    End Sub

Private Sub UserControl_Initialize()
    lblStatus.Caption = "Status: Idle."
    End Sub
    

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With ftpControlVals
        .intPort = PropBag.ReadProperty("ftp_remote_port", "21")
        .strAddress = PropBag.ReadProperty("ftp_remote_address", "ftp://ftp.")
        .strLocalFilename = PropBag.ReadProperty("ftp_local_filename", "c:\")
        .strRemoteFilename = PropBag.ReadProperty("ftp_remote_filename", "/public_html/")
        .boolAsciiMode = PropBag.ReadProperty("ftp_is_ascii_mode", False)
        .strPassword = PropBag.ReadProperty("ftp_password", "")
        .strUsername = PropBag.ReadProperty("ftp_username", "anonymous")
        BackColor = PropBag.ReadProperty("BackColour", "&H8000000F&") ' grey
        lblStatus.ForeColor = PropBag.ReadProperty("FontColour", "&H80000012&") ' blk
        cmdCancel.Caption = PropBag.ReadProperty("CancelButton_Text", "Cancel Transfer")
        cmdCancel.Enabled = PropBag.ReadProperty("CancelButton_Enabled", True)
        cmdCancel.Width = PropBag.ReadProperty("CancelButton_Width", "1575")
    End With
    End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With pbTransferStatus
        .Width = Width - (.Left * 2)
    End With
    
    With lblStatus
        .Width = Width - (.Left * 2)
    End With
    
    With cmdCancel
        .Left = (Width - .Width) / 2
    End With
    
    Height = 1410
    End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ftp_remote_port", ftpControlVals.intPort, "21")
    Call PropBag.WriteProperty("ftp_remote_address", ftpControlVals.strAddress, "ftp://ftp.")
    Call PropBag.WriteProperty("ftp_local_filename", ftpControlVals.strLocalFilename, "c:\")
    Call PropBag.WriteProperty("ftp_remote_filename", ftpControlVals.strRemoteFilename, "/public_html/")
    Call PropBag.WriteProperty("ftp_password", ftpControlVals.strPassword, "me@mee.com")
    Call PropBag.WriteProperty("ftp_username", ftpControlVals.strUsername, "anonymous")
    Call PropBag.WriteProperty("ftp_is_ascii_mode", ftpControlVals.boolAsciiMode, False)
    Call PropBag.WriteProperty("BackColour", BackColor, "&H8000000F&") ' grey
    Call PropBag.WriteProperty("FontColour", lblStatus.ForeColor, "&H80000012&") ' blk
    Call PropBag.WriteProperty("CancelButton_Text", cmdCancel.Caption, "Cancel Transfer")
    Call PropBag.WriteProperty("CancelButton_Enabled", cmdCancel.Enabled, True)
    Call PropBag.WriteProperty("CancelButton_Width", cmdCancel.Width, "1575")
    End Sub

Public Property Get ftp_remote_address() As String
Attribute ftp_remote_address.VB_Description = "sets remote ftp address"
    ftp_remote_address = ftpControlVals.strAddress
    End Property

Public Property Let ftp_remote_address(ByVal vNewValue As String)
    ftpControlVals.strAddress = vNewValue
    PropertyChanged "ftp_remote_address"
    End Property

Public Property Get ftp_password() As String
Attribute ftp_password.VB_Description = "password to use when connecting"
    ftp_password = ftpControlVals.strPassword
    End Property

Public Property Let ftp_password(ByVal vNewValue As String)
    ftpControlVals.strPassword = vNewValue
    PropertyChanged "ftp_password"
    End Property

Public Property Get ftp_is_ascii_mode() As Boolean
Attribute ftp_is_ascii_mode.VB_Description = "Set's transfer mode to ASCII or Binary. Set to ASCII for ALL text files, and Binary for non-text (default)"
    ftp_is_ascii_mode = ftpControlVals.boolAsciiMode
    End Property

Public Property Let ftp_is_ascii_mode(ByVal vNewValue As Boolean)
    ftpControlVals.boolAsciiMode = vNewValue
    PropertyChanged "ftp_is_ascii_mode"
    End Property

Public Property Get ftp_username() As String
Attribute ftp_username.VB_Description = "sets the username to use when connecting"
    ftp_username = ftpControlVals.strUsername
    End Property

Public Property Let ftp_username(ByVal vNewValue As String)
    ftpControlVals.strUsername = vNewValue
    PropertyChanged "ftp_username"
    End Property

Public Property Get ftp_local_filename() As String
Attribute ftp_local_filename.VB_Description = "Local filename for upload/downloaded file."
    ftp_local_filename = ftpControlVals.strLocalFilename
    End Property

Public Property Let ftp_local_filename(ByVal vNewValue As String)
    ftpControlVals.strLocalFilename = vNewValue
    PropertyChanged "ftp_local_filename"
    End Property

Public Property Get ftp_remote_filename() As String
Attribute ftp_remote_filename.VB_Description = "remote filename to download (use full path: /public_html/my_docs/file1.html)"
    ftp_remote_filename = ftpControlVals.strRemoteFilename
    End Property

Public Property Let ftp_remote_filename(ByVal vNewValue As String)
    ftpControlVals.strRemoteFilename = vNewValue
    PropertyChanged "ftp_remote_filename"
    End Property

Public Property Get BackColour() As OLE_COLOR
    BackColour = BackColor
    End Property

Public Property Let BackColour(ByVal vNewValue As OLE_COLOR)
    BackColor = vNewValue
    PropertyChanged "BackColour"
    End Property

Public Property Get FontColour() As OLE_COLOR
    FontColour = lblStatus.ForeColor
    End Property

Public Property Let FontColour(ByVal vNewValue As OLE_COLOR)
    lblStatus.ForeColor = vNewValue
    PropertyChanged "FontColour"
    End Property

Public Property Get CancelButton_Text() As String
    CancelButton_Text = cmdCancel.Caption
    End Property

Public Property Let CancelButton_Text(ByVal vNewValue As String)
    cmdCancel.Caption = vNewValue
    PropertyChanged "CancelButton_Text"
    End Property
    

Public Property Get CancelButton_Enabled() As Boolean
    CancelButton_Enabled = cmdCancel.Enabled
    End Property

Public Property Let CancelButton_Enabled(ByVal vNewValue As Boolean)
    cmdCancel.Enabled = vNewValue
    PropertyChanged "CancelButton_Enabled"
    End Property

Public Property Get CancelButton_Width() As Long
    CancelButton_Width = cmdCancel.Width
    End Property

Public Property Let CancelButton_Width(ByVal vNewValue As Long)
    cmdCancel.Width = vNewValue
    PropertyChanged "CancelButton_Width"
    End Property


'**************************************************************
Public Function ftp_Download_quick(ByRef ReturnResult As String) As Boolean
Attribute ftp_Download_quick.VB_Description = "Downloads a file with no events raised. Use this for small files."
'**************************************************************
    ' Start Download
    ftp_Download_quick = False
    ReturnResult = "Unknown"
    
    
    Dim ftpFile As ftpCurrentXfer
    With ftpControlVals
        ftpFile.boolAsciiMode = .boolAsciiMode
        ftpFile.intPort = .intPort
        ftpFile.strLocalFilename = .strLocalFilename
        ftpFile.strPassword = .strPassword
        ftpFile.strRemoteFilename = .strRemoteFilename
        ftpFile.strUsername = .strUsername
        ftpFile.strAddress = .strAddress
        ftpFile.dblFileLoc = 0
        ftpFile.dblFileSize = 0
        ftpFile.intFileNo = 0
    End With
    
    ' Connect to FTP Site, check username/password etc.
    Dim ftp_sessID As Long
    Dim hFile      As Long
    Dim Service  As Long
    Dim InetSess As Long
 
    ' Create Session.
    UpdateStatus 0, 0, "Creating Session.."
    InetSess = InternetOpen("MYOW QuickFTP Download Session", INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    If InetSess = 0 Then
        ReturnResult = "Error creating InetSession. Please check your Internet Explorer Version and Internet connection status (working online/offline)"
        UpdateStatus 1, 4, ReturnResult
        ftp_Download_quick = False
        InternetCloseHandle InetSess
        Exit Function
        End If
    
 
    ' Create a FTP Session
    UpdateStatus 0, 0, "Creating FTP Session.."
    ftp_sessID = InternetConnect(InetSess, ftpFile.strAddress, ftpFile.intPort, ftpFile.strUsername, ftpFile.strPassword, INTERNET_SERVICE_FTP, Service, &H0)
    If ftp_sessID = 0 Then
        ReturnResult = "Unable to connect. Please check your username, password, FTP Address and Port. Alternativly the server may be down."
        UpdateStatus 2, 4, ReturnResult
        ftp_Download_quick = False
        InternetCloseHandle ftp_sessID
        InternetCloseHandle InetSess
        Exit Function
    End If
 
    UpdateStatus 3, 4, "Connected! Downloading File..  (" & IIf(ftpFile.boolAsciiMode, "Ascii Mode", "Binary Mode") & ").. Please Wait."
   
    'Download the file the quick way.
    If FtpGetFile(ftp_sessID, ftpFile.strRemoteFilename & vbNullChar, ftpFile.strLocalFilename & vbNullChar, 0&, 0&, IIf(ftpFile.boolAsciiMode = True, FTP_TRANSFER_TYPE_ASCII, FTP_TRANSFER_TYPE_BINARY), 0&) = 0 Then
        UpdateStatus 4, 4, "Failed"
        ReturnResult = "Failed downloading file, check file permissions (both remote & local) and try again."
        ftp_Download_quick = False
    Else
        UpdateStatus 4, 4, "Download Success, disconnecting."
        ftp_Download_quick = True
        ReturnResult = "download complete."
    End If
    
 
    ' Cleanup.
    InternetCloseHandle ftp_sessID
    InternetCloseHandle InetSess
    
    End Function


'**************************************************************
Public Function ftp_Upload_quick(ByRef ReturnResult As String) As Boolean
'**************************************************************
    ' Start Download
    
    ftp_Upload_quick = False
    ReturnResult = "Unknown"
    
    
    Dim ftpFile As ftpCurrentXfer
    With ftpControlVals
        ftpFile.boolAsciiMode = .boolAsciiMode
        ftpFile.intPort = .intPort
        ftpFile.strLocalFilename = .strLocalFilename
        ftpFile.strPassword = .strPassword
        ftpFile.strRemoteFilename = .strRemoteFilename
        ftpFile.strUsername = .strUsername
        ftpFile.strAddress = .strAddress
        ftpFile.dblFileLoc = 0
        ftpFile.dblFileSize = 0
        ftpFile.intFileNo = 0
    End With
    
    ' Connect to FTP Site, check username/password etc.
    Dim ftp_sessID As Long
    Dim hFile      As Long
    Dim Service  As Long
    Dim InetSess As Long
 
    ' Create Session.
    UpdateStatus 1, 4, "Creating Session.."
    InetSess = InternetOpen("MYOW QuickFTP Upload Session", INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    If InetSess = 0 Then
        ReturnResult = "Error creating InetSession. Please check your Internet Explorer Version and Internet connection status (working online/offline)"
        UpdateStatus 4, 4, ReturnResult
        ftp_Upload_quick = False
        InternetCloseHandle InetSess
        Exit Function
        End If
    
 
    ' Create a FTP Session
    UpdateStatus 2, 4, "Creating FTP Session.."
    ftp_sessID = InternetConnect(InetSess, ftpFile.strAddress, ftpFile.intPort, ftpFile.strUsername, ftpFile.strPassword, INTERNET_SERVICE_FTP, Service, &H0)
    If ftp_sessID = 0 Then
        ReturnResult = "Unable to connect. Please check your username, password, FTP Address and Port. Alternativly the server may be down."
        UpdateStatus 4, 4, ReturnResult
        ftp_Upload_quick = False
        InternetCloseHandle ftp_sessID
        InternetCloseHandle InetSess
        Exit Function
    End If
 
    UpdateStatus 3, 4, "Connected! Uploading File..  (" & IIf(ftpFile.boolAsciiMode, "Ascii Mode", "Binary Mode") & ").. Please Wait."
   
    'Download the file the quick way.
    If FtpPutFile(ftp_sessID, ftpFile.strLocalFilename & vbNullChar, ftpFile.strRemoteFilename & vbNullChar, IIf(ftpFile.boolAsciiMode = True, FTP_TRANSFER_TYPE_ASCII, FTP_TRANSFER_TYPE_BINARY), 0&) = 0 Then
        UpdateStatus 4, 4, "Failed"
        ReturnResult = "Failed uploading file, check file permissions (both remote & local) and try again."
        ftp_Upload_quick = False
    Else
        UpdateStatus 4, 4, "Upload Success, disconnecting."
        ftp_Upload_quick = True
        ReturnResult = "upload complete."
    End If
    
 
    ' Cleanup.
    InternetCloseHandle ftp_sessID
    InternetCloseHandle InetSess
    
    End Function

