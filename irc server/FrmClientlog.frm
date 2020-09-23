VERSION 5.00
Begin VB.Form FrmClientlog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clients log"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Lstlog 
      Height          =   2400
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Lblip 
      Caption         =   "ip"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lblnickname 
      Caption         =   "nick"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblwinsocknum 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Lbljoin 
      Caption         =   "join"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Lblpart 
      Caption         =   "part"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "FrmClientlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Unload Me
End Sub
