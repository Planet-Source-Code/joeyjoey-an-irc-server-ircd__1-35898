VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "noIRCD"
   ClientHeight    =   3015
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4530
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNickname 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtServer 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "noIRC.Serv"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Timer TmrUpload 
      Interval        =   100
      Left            =   2400
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   2040
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6667
   End
   Begin VB.Label LblEli 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uptime : 00:00:00"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Image ImgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   3060
      Left            =   0
      Picture         =   "FrmMain.frx":0000
      Top             =   0
      Width           =   4560
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuConfigure 
      Caption         =   "&Configure"
      Begin VB.Menu mnuOnCon 
         Caption         =   "&On Connect"
      End
   End
   Begin VB.Menu mnuRehash 
      Caption         =   "&Rehash"
      Begin VB.Menu mnuonconerefresh 
         Caption         =   "O&n Connect"
      End
   End
   Begin VB.Menu mnuLOGS 
      Caption         =   "&Logs"
      Begin VB.Menu mnuClients 
         Caption         =   "C&lients"
      End
      Begin VB.Menu mnuServer 
         Caption         =   "&Server"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vstart As Variant
Dim vendtime As Variant
Dim velapsed As Variant
Dim vsock As Integer
Dim vmyip As String
Dim vserv As String
Dim vurip As String
Dim vsele As String
Dim vloads As Integer
Dim vstr As String
Dim vste As String


Private Sub Form_Load()
vstart = Now
'vnums = 0
vsock = 1
Load WS(1)
WS(vsock).Listen
End Sub



Private Sub mnuClients_Click()
FrmClientlog.Show
End Sub

Private Sub mnuexit_Click()
FrmQuit.Show
End Sub

Private Sub mnuOnCon_Click()
FrmOnConnect.Show
End Sub

Private Sub mnuonconerefresh_Click()
Unload FrmOnConnect
End Sub

Private Sub TmrUpload_Timer()
vendtime = Now
velapsed = vendtime - vstart
LblEli = "Uptime : " & Format(velapsed, "hh:mm:ss")
End Sub

Private Sub TxtServer_Change()
Me.Caption = "noIRC - " & TxtServer
End Sub

Private Sub WS_Close(Index As Integer)
WS(Index).Close
If WS(Index).State <> 0 Then WS(Index).Listen
FrmClientlog.Lstlog.AddItem ("Quit " & WS(Index).Index & " " & WS(Index).RemoteHostIP & " " & Now)
'FrmClientlog.Lblpart(Index).Caption = Now
Unload WS(Index)

End Sub

Private Sub WS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If WS(Index).State <> 0 Then WS(Index).Close
WS(Index).Accept requestID
vsock = vsock + 1
Load WS(vsock)
WS(vsock).LocalPort = 6667
WS(vsock).Listen

vmyip = Replace(FrmOnConnect.Txtmsgsent, "$myip", FrmMain.WS(0).LocalIP)
FrmOnConnect.Txtmsgsent = vmyip
vserv = Replace(FrmOnConnect.Txtmsgsent, "$serv", FrmMain.TxtServer)
FrmOnConnect.Txtmsgsent = vserv
vurip = Replace(FrmOnConnect.Txtmsgsent, "$urip", FrmMain.WS(Index).RemoteHostIP)
FrmOnConnect.Txtmsgsent = vurip

WS(Index).SendData (FrmOnConnect.Txtmsgsent)
End Sub

Private Sub WS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Call WS(Index).GetData(strData, vbString)
If strData Like ":*NICK*" Then
End If
If strData Like "NICK*" Then
    TxtNickname = ""
    TxtNickname = Mid(strData, 6, Len(strData))

    vsele = 0
    Do Until TxtNickname.SelText = Chr(10)
        TxtNickname.SelStart = vsele
        TxtNickname.SelLength = 1
        vstr = vstr & TxtNickname.SelText
        vsele = vsele + 1
    Loop
    vste = Mid(vstr, 1, Len(vstr) - 1)
    TxtNickname = vste
     FrmClientlog.Lstlog.AddItem (WS(Index).Index & " " & vste & " " & WS(Index).RemoteHostIP & " " & Now)
vloads = vloads + 1
Load FrmClientlog.Lblip(vloads)
Load FrmClientlog.Lblnickname(vloads)
Load FrmClientlog.lblwinsocknum(vloads)
Load FrmClientlog.Lbljoin(vloads)
Load FrmClientlog.Lblpart(vloads)

FrmClientlog.Lblip(vloads).Top = FrmClientlog.Lblip(vloads - 1).Top + 255
FrmClientlog.Lblnickname(vloads).Top = FrmClientlog.Lblnickname(vloads - 1).Top + 255
FrmClientlog.lblwinsocknum(vloads).Top = FrmClientlog.lblwinsocknum(vloads - 1).Top + 255
FrmClientlog.Lbljoin(vloads).Top = FrmClientlog.Lbljoin(vloads - 1).Top + 255
FrmClientlog.Lblpart(vloads).Top = FrmClientlog.Lblpart(vloads - 1).Top + 255


'FrmClientlog.Lblip(vloads).Visible = True
'FrmClientlog.Lblnickname(vloads).Visible = True
'FrmClientlog.lblwinsocknum(vloads).Visible = True
'FrmClientlog.Lbljoin(vloads).Visible = True
'FrmClientlog.Lblpart(vloads).Visible = True

FrmClientlog.Lblip(vloads).Caption = WS(Index).RemoteHostIP
FrmClientlog.Lblnickname(vloads).Caption = vste
FrmClientlog.lblwinsocknum(vloads).Caption = WS(Index).Index
FrmClientlog.Lbljoin(vloads).Caption = Now
FrmClientlog.Lblpart(vloads).Caption = "N/A"

TxtNickname = ""
vste = ""
vstr = ""
End If
End Sub
