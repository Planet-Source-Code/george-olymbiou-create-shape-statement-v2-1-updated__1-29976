VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create SHAPE Statement From George Olymbiou"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAntiSelect2 
      Caption         =   "Undo Select All"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdAntiSelect1 
      Caption         =   "Undo Select All"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdSelectAllDetail 
      Caption         =   "Select All"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelectAllMaster 
      Caption         =   "Select All"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
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
      Left            =   1883
      TabIndex        =   12
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton CmdClose 
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
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   5895
   End
   Begin VB.ComboBox JoinFieldsOfDetailTable 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5280
      Width           =   2775
   End
   Begin VB.ComboBox JoinFieldsOfMasterTable 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5280
      Width           =   2775
   End
   Begin VB.ListBox FieldsOfDetailTable 
      Height          =   2310
      Left            =   3120
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ListBox FieldsOfMasterTable 
      Height          =   2310
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox DetailTable 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.ComboBox MasterTable 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2775
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
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   5040
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   19
      Top             =   2040
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fields Of Detail Table"
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
      Left            =   3480
      TabIndex        =   16
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Join Fields"
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
      TabIndex        =   18
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fields Of Master Table"
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
      TabIndex        =   17
      Top             =   1800
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Detail :"
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
      Left            =   3120
      TabIndex        =   15
      Top             =   360
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Master :"
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
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Database :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Is A SHAPE Generator Made
'By George Olymbiou
'Date 20/12/2001
Dim WithEvents Conn As ADODB.Connection
Attribute Conn.VB_VarHelpID = -1
Public Sub ResetAll()
On Error GoTo ErrDescr
    MasterTable.Clear
    DetailTable.Clear
    FieldsOfMasterTable.Clear
    FieldsOfDetailTable.Clear
    JoinFieldsOfMasterTable.Clear
    JoinFieldsOfDetailTable.Clear
Exit Sub
ErrDescr:
MsgBox Err.Description
End Sub
Private Sub cmdAntiSelect1_Click()
On Error GoTo ErrDescr
    If FieldsOfMasterTable.ListCount <> 0 Then
        For i = 0 To FieldsOfMasterTable.ListCount - 1
            FieldsOfMasterTable.Selected(i) = False
        Next
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdAntiSelect2_Click()
On Error GoTo ErrDescr
    If FieldsOfDetailTable.ListCount <> 0 Then
        For i = 0 To FieldsOfDetailTable.ListCount - 1
            FieldsOfDetailTable.Selected(i) = False
        Next
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdBrowse_Click()
On Error GoTo ErrDescr
    Form3.Show
Exit Sub
ErrDescr:
        MsgBox Err.Description
End Sub
Private Sub CmdClose_Click()
On Error GoTo ErrDescr
    End
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdOK_Click()
On Error GoTo ErrDescr
    If IsOK(Form1) = True Then
        Dim ParentCMDFields As String
        Dim ChildCMDFields As String
        Dim FinalSHAPEStatement As String
            
            ParentCMDFields = ParentCMDFields + "[" + JoinFieldsOfMasterTable.List(0) + "]"
            
            For i = 1 To JoinFieldsOfMasterTable.ListCount - 1
                ParentCMDFields = ParentCMDFields + "," + "[" + JoinFieldsOfMasterTable.List(i) + "]"
            Next
                        
            ChildCMDFields = ChildCMDFields + "[" + JoinFieldsOfDetailTable.List(0) + "]"
            
            For i = 1 To JoinFieldsOfDetailTable.ListCount - 1
                ChildCMDFields = ChildCMDFields + "," + "[" + JoinFieldsOfDetailTable.List(i) + "]"
            Next

            FinalSHAPEStatement = "SHAPE {SELECT " + ParentCMDFields + " FROM [" + MasterTable.Text + "] } AS ParentCMD APPEND ({SELECT " + ChildCMDFields + " FROM [" + DetailTable.Text + "] } AS ChildCMD RELATE [" + JoinFieldsOfMasterTable.Text + "] TO [" + JoinFieldsOfDetailTable.Text + "] ) AS ChildCMD "
            Form2.Show
            Form2.RichTextBox1.Text = FinalSHAPEStatement
            Form2.RichTextBox2.Text = Module1.ConnectionString
    Else
        MsgBox "You Must Suply All Information"
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdSelectAllDetail_Click()
On Error GoTo ErrDescr
    If FieldsOfDetailTable.ListCount <> 0 Then
        For i = 0 To FieldsOfDetailTable.ListCount - 1
            FieldsOfDetailTable.Selected(i) = True
        Next
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdSelectAllMaster_Click()
On Error GoTo ErrDescr
    If FieldsOfMasterTable.ListCount <> 0 Then
        For i = 0 To FieldsOfMasterTable.ListCount - 1
            FieldsOfMasterTable.Selected(i) = True
        Next
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub DetailTable_Click()
On Error GoTo ErrDescr
    If DetailTable.Text <> "" Then
        FieldsOfDetailTable.Clear
            Conn.Open ConnectionString
                Dim MyCatalog As New ADOX.Catalog
                Dim MyTable As ADOX.Table
                Set MyCatalog.ActiveConnection = Conn
                Set MyTable = MyCatalog.Tables(DetailTable.Text)
        
                    For i = 0 To MyTable.Columns.Count - 1
                        FieldsOfDetailTable.AddItem MyTable.Columns(i).Name
                    Next

                Set MyCatalog = Nothing
                Set MyTable = Nothing
            Conn.Close
    JoinFieldsOfDetailTable.Clear
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub FieldsOfDetailTable_Click()
On Error GoTo ErrDescr
    JoinFieldsOfDetailTable.Clear
        For i = 0 To FieldsOfDetailTable.ListCount - 1
            If FieldsOfDetailTable.Selected(i) = True Then
                JoinFieldsOfDetailTable.AddItem FieldsOfDetailTable.List(i)
            End If
        Next
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub MasterTable_Click()
On Error GoTo ErrDescr
    If MasterTable.Text <> "" Then
        FieldsOfMasterTable.Clear
            Conn.Open ConnectionString
                Dim MyCatalog As New ADOX.Catalog
                Dim MyTable As ADOX.Table
                    Set MyCatalog.ActiveConnection = Conn
                    Set MyTable = MyCatalog.Tables(MasterTable.Text)
 
                        For i = 0 To MyTable.Columns.Count - 1
                            FieldsOfMasterTable.AddItem MyTable.Columns(i).Name
                        Next

                    Set MyCatalog = Nothing
                    Set MyTable = Nothing
            Conn.Close
        JoinFieldsOfMasterTable.Clear
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub FieldsOfMasterTable_Click()
On Error GoTo ErrDescr
    JoinFieldsOfMasterTable.Clear
        For i = 0 To FieldsOfMasterTable.ListCount - 1
            If FieldsOfMasterTable.Selected(i) = True Then
                JoinFieldsOfMasterTable.AddItem FieldsOfMasterTable.List(i)
            End If
        Next
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub Form_Load()
On Error GoTo ErrDescr
    Set Conn = New ADODB.Connection
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrDescr
    Set Conn = Nothing
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub


