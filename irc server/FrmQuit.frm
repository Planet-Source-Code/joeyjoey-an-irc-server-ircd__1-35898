VERSION 5.00
Begin VB.Form FrmQuit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quit?"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdYes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton CmdNo 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label LblEli 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Lblsure 
      Alignment       =   2  'Center
      Caption         =   "If you quit all connected clients will get disconnected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   100
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "FrmQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNo_Click()
Unload Me
End Sub

Private Sub CmdYes_Click()
Unload Me
End
End Sub
