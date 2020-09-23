VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SHAPE"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdtestSHAPE 
      Caption         =   "Test SHAPE Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   5775
   End
   Begin VB.CommandButton cmdCopyToClipboard2 
      Caption         =   "Copy To Clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdCopyToClipBoard1 
      Caption         =   "Copy To Clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form2.frx":0000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   5775
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form2.frx":00DF
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Connection String :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SHAPE Statement :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopyToClipBoard1_Click()
    Clipboard.Clear
    Clipboard.SetText RichTextBox1.Text
End Sub

Private Sub cmdCopyToClipboard2_Click()
    Clipboard.Clear
    Clipboard.SetText RichTextBox2.Text
End Sub

Private Sub cmdtestSHAPE_Click()
On Error GoTo ErrDescr
    Form4.Show
    Form4.OpenConnection RichTextBox2.Text
    Form4.OpenRecordset RichTextBox1.Text
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    RichTextBox1.AutoVerbMenu = True
    RichTextBox2.AutoVerbMenu = True
    RichTextBox1.Locked = True
    RichTextBox2.Locked = True
End Sub
