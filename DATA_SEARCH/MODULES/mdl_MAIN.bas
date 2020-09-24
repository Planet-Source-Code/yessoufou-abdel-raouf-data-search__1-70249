Attribute VB_Name = "mdl_MAIN"


Public Sub Main()

    Call fn_OPEN_CONNECTION
    
    strPicturePath = App.Path & "/PICTURES/"
    
    frm_MAIN.Show
    frm_EMPLOYEES_SEARCH.Show
    
End Sub



Public Function fn_OPEN_CONNECTION() As ADODB.Connection
On Error GoTo errHandler

    Dim con_Obj As New ADODB.Connection
    
    With SystemData
        con_Obj.ConnectionString = "provider = microsoft.jet.oledb.4.0 ; data source = " & App.Path & "\Database\DATA_SEARCH.mdb;Jet OLEDB:Database Password=GODISGR8"
    End With

    con_Obj.Open
    Set fn_OPEN_CONNECTION = con_Obj

    
EXITPROCEDURE:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "ClsDatabase", "DBConnection")
    GoTo EXITPROCEDURE
    
End Function

Public Function fn_CLOSE_CONNECTION(ByRef con_Obj As ADODB.Connection)
On Error GoTo errHandler
    
    con_Obj.Close
    Set con_Obj = Nothing
    
EXITPROCEDURE:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "ClsDatabase", "DBConnection")
    GoTo EXITPROCEDURE
    
End Function
