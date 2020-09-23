Attribute VB_Name = "Module1"
Public ConnectionString As String
Public Sub BuildConnectionString(IsAccess As Boolean, Datasource, UserName, Password, Server As String)
    If IsAccess = True Then
        ConnectionString = "PROVIDER=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Datasource + ";Persist Security Info=False;User ID=" + UserName + ";Password=" + Password + ";"
    Else
        ConnectionString = "PROVIDER=MSDataShape;Data Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + UserName + ";Initial Catalog=" + Server + ";Data Source=" + Datasource + ";Password=" + Password + ";"
    End If
End Sub
Public Function IsOK(myform As Form) As Boolean
IsOK = True
    If myform.MasterTable.Text = myform.DetailTable.Text Then
        IsOK = False
        Exit Function
    End If
    If myform.MasterTable.Text = "" Then
        IsOK = False
        Exit Function
    End If
    If myform.DetailTable.Text = "" Then
        IsOK = False
        Exit Function
    End If
    If myform.FieldsOfMasterTable.ListCount = 0 Then
        IsOK = False
        Exit Function
    End If
    If myform.FieldsOfDetailTable.ListCount = 0 Then
        IsOK = False
        Exit Function
    End If
    If myform.JoinFieldsOfMasterTable.Text = "" Then
        IsOK = False
        Exit Function
    End If
    If myform.JoinFieldsOfDetailTable.Text = "" Then
        IsOK = False
        Exit Function
    End If
    
    Dim MyBool1, MyBool2 As Boolean
    
    MyBool1 = False
    
    For i = 0 To myform.FieldsOfMasterTable.ListCount - 1
        If myform.FieldsOfMasterTable.Selected(i) = True Then
            MyBool1 = True
        End If
    Next
    
    MyBool2 = False
    
    For i = 0 To myform.FieldsOfDetailTable.ListCount - 1
        If myform.FieldsOfDetailTable.Selected(i) = True Then
            MyBool2 = True
        End If
    Next
    
    If MyBool1 = False Or MyBool2 = False Then
        IsOK = False
    End If
    
End Function

