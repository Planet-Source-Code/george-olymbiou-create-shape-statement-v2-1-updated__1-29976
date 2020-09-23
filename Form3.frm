VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect To Database"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboServer 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   5775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   1620
      TabIndex        =   7
      Top             =   2760
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server :"
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
      Index           =   4
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   825
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   2
      Left            =   120
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   1560
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   120
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Driver :"
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
      Index           =   0
      Left            =   750
      TabIndex        =   11
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database :"
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
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password :"
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
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name :"
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
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   720
      Width           =   1305
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
        CDialog.DialogTitle = "Select An .mdb File"
        CDialog.FileName = "*.mdb"
        CDialog.Filter = "*.mdb"
        CDialog.ShowOpen
        txtDatabase.Text = CDialog.FileName
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
On Error GoTo ErrDescr

Form1.ResetAll

If Combo1.Text = "" Then
    MsgBox "You Must Suply A Driver To Continue", vbCritical
    Exit Sub
End If

If Combo1.Text = "Microsoft Access Driver (*.mdb)" Then
    If txtDatabase.Text = "" Then
        MsgBox "You Must Suply An Access Database", vbCritical
        Exit Sub
    End If
ElseIf Combo1.Text = "SQL Server" Then
    If txtDatabase.Text = "" Then
        MsgBox "You Must Suply A Database", vbCritical
        Exit Sub
    ElseIf cboServer.Text = "" Then
        MsgBox "You Must Suply A Server", vbCritical
        Exit Sub
    ElseIf txtUserName.Text = "" Then
        MsgBox "You Must Suply A UserName", vbCritical
        Exit Sub
    End If
End If

Dim fso As New Scripting.FileSystemObject

    If Combo1.Text = "Microsoft Access Driver (*.mdb)" Then
        If fso.FileExists(txtDatabase.Text) = False Or fso.GetExtensionName(txtDatabase.Text) <> "mdb" Then
            MsgBox "This File Is Not An Access File"
            Exit Sub
        End If
        
        Module1.BuildConnectionString True, txtDatabase.Text, txtUserName.Text, txtPassword.Text, ""
    ElseIf Combo1.Text = "SQL Server" Then
        Module1.BuildConnectionString False, txtDatabase.Text, txtUserName.Text, txtPassword.Text, cboServer.Text
    End If

    Dim Conn As New ADODB.Connection
    Dim MyCatalog As New ADOX.Catalog
    Dim AllTables As ADOX.Tables
    
    Conn.Open Module1.ConnectionString
        
        Set MyCatalog.ActiveConnection = Conn
        Set AllTables = MyCatalog.Tables
        
        For i = 0 To AllTables.Count - 1
            If AllTables.Item(i).Type <> "SYSTEM TABLE" And AllTables.Item(i).Type <> "GLOBAL TEMPORARY" And AllTables.Item(i).Type <> "ACCESS TABLE" Then
                Form1.MasterTable.AddItem AllTables.Item(i).Name
                Form1.DetailTable.AddItem AllTables.Item(i).Name
            End If
        Next
        
    Conn.Close
    
    Set AllTables = Nothing
    Set MyCatalog = Nothing
    
    Set Conn = Nothing
    
    Unload Me
    
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub Combo1_Click()
On Error GoTo ErrDescr
    If Combo1.Text = "Microsoft Access Driver (*.mdb)" Then
        cboServer.Enabled = False
        cmdBrowse.Enabled = True
        txtDatabase.Locked = True
    End If
    If Combo1.Text = "SQL Server" Then
        txtDatabase.Locked = False
        cboServer.Enabled = True
        cmdBrowse.Enabled = False
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub Form_Load()
On Error GoTo ErrDescr
    txtDatabase.Locked = True
    cmdBrowse.Enabled = False
    cboServer.Enabled = False
    Combo1.AddItem "Microsoft Access Driver (*.mdb)"
    Combo1.AddItem "SQL Server"
    cboServer.AddItem "Local"
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
