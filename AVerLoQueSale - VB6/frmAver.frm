VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Listitems"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   4815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Text            =   "c:\Temp"
         Top             =   0
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FindPreparando
FindFiles Text1.Text


End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
Set fs = New FileSystemObject
End Sub

Private Sub Form_Resize()
lv1.Move lv1.Left, 0, Me.ScaleWidth - lv1.Left, Me.ScaleHeight - Command1.Height - 140
Frame1.Top = lv1.Height + 140
End Sub

Private Sub lv1_DblClick()
'Shell "explorer.exe " & lv1.SelectedItem.SubItems(1) & "\" & lv1.SelectedItem.Text, vbNormalFocus

ShellExecute Me.hwnd, "open", lv1.SelectedItem.SubItems(1) & "\" & lv1.SelectedItem.Text, "", "", SW_NORMAL

End Sub
