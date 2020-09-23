VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SHAPE TEST"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   5895
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DGridDetail 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGridMaster 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2778
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   6135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Conn As ADODB.Connection
Attribute Conn.VB_VarHelpID = -1
Public WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdMoveFirst_Click()
On Error GoTo ErrDescr
    rst.MoveFirst
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdMoveLast_Click()
On Error GoTo ErrDescr
    rst.MoveLast
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdMoveNext_Click()
On Error GoTo ErrDescr
    If rst.AbsolutePosition <> rst.RecordCount Then
        rst.MoveNext
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub cmdMovePrevious_Click()
On Error GoTo ErrDescr
    If rst.AbsolutePosition <> 1 Then
        rst.MovePrevious
    End If
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Private Sub Form_Load()
On Error GoTo ErrDescr
    Set Conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    DGridMaster.AllowUpdate = False
    DGridDetail.AllowUpdate = False
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
Public Sub OpenConnection(ConnectionString As String)
On Error GoTo ErrDescr
    Conn.CursorLocation = adUseClient
    Conn.Open ConnectionString
Exit Sub
ErrDescr:
    MsgBox "You Have A Mistake In Your Connection String, Correct It And Try Again", vbCritical, "Error Message"
End Sub
Public Sub OpenRecordset(SqlStatement As String)
On Error GoTo ErrDescr
    rst.Open SqlStatement, Conn, 3, 3
    If rst.RecordCount > 0 Then
        Set DGridMaster.Datasource = rst
        Set DGridDetail.Datasource = rst.Fields("childcmd").UnderlyingValue
    Else
        MsgBox "Your Shape Statement Must Have A Record In Master Table", vbCritical, "Error Message"
        DGridMaster.Enabled = False
        DGridDetail.Enabled = False
        cmdMoveFirst.Enabled = False
        cmdMoveLast.Enabled = False
        cmdMoveNext.Enabled = False
        cmdMovePrevious.Enabled = False
    End If
Exit Sub
ErrDescr:
    MsgBox "You Have A Mistake In Your SHAPE Statement, Correct It And Try Again", vbCritical, "Error Message"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrDescr
    If rst.State = adStateOpen Then
        rst.Close
    End If
    If Conn.State = adStateOpen Then
        Conn.Close
    End If
    Set rst = Nothing
    Set Conn = Nothing
Exit Sub
ErrDescr:
    MsgBox Err.Description
End Sub
