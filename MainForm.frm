VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Converter"
   ClientHeight    =   2028
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6048
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2028
   ScaleWidth      =   6048
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5280
      TabIndex        =   4
      Top             =   960
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5280
      TabIndex        =   3
      Top             =   360
      Width           =   492
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   600
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt/Decrypt"
      Height          =   372
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
      Width           =   1812
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5052
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5052
   End
   Begin VB.Label Label2 
      Caption         =   "Destination File"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5052
   End
   Begin VB.Label Label1 
      Caption         =   "Source File"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5052
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then Exit Sub
Command1.Caption = "Working"
MainForm.Enabled = False
Screen.MousePointer = vbHourglass
Call ConvertFile(Text1.Text, Text2.Text)
Screen.MousePointer = vbDefault
MainForm.Enabled = True
Command1.Caption = "Encrypt/Decrypt"
End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "All Files|*.*"
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
CommonDialog1.InitDir = CommonDialog1.FileName
End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "All Files|*.*"
CommonDialog1.ShowSave
Text2.Text = CommonDialog1.FileName
CommonDialog1.InitDir = CommonDialog1.FileName
End Sub
