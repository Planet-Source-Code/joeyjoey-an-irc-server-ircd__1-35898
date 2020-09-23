VERSION 5.00
Begin VB.Form FrmOnConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "On Connect"
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
   Begin VB.CommandButton CmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Txtmsgsent 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmOnConnect.frx":0000
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmOnConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Me.Hide

End Sub
