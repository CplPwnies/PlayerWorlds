VERSION 5.00
Begin VB.Form frmMapKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Key"
   ClientHeight    =   2310
   ClientLeft      =   105
   ClientTop       =   210
   ClientWidth     =   4815
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMapKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTake 
      Caption         =   "Take key away upon use"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.HScrollBar scrlItem 
      Height          =   255
      Left            =   840
      Max             =   500
      Min             =   1
      TabIndex        =   2
      Top             =   600
      Value           =   1
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmMapKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblName.Caption = Trim$(Item(scrlItem.Value).name)
End Sub

Private Sub cmdOK_Click()
    KeyEditorNum = scrlItem.Value
    KeyEditorTake = chkTake.Value
    Unload Me
End Sub

Private Sub scrlItem_Change()
    lblItem.Caption = STR(scrlItem.Value)
    lblName.Caption = Trim$(Item(scrlItem.Value).name)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

