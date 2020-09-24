VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_EMPLOYEES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim con_Obj As New ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarEmployeeID As Long 'local copy
Private mvarLastName As String 'local copy
Private mvarFirstName As String 'local copy
Private mvarTitleOfCourtesy As String 'local copy
Private mvarAddress As String 'local copy
Private mvarCountry As String 'local copy
Private mvarHomePhone As String 'local copy
Private mvarPhoto As String 'local copy
Private mvarNotes As String 'local copy
Private mvarWorkingStatus As Long 'local copy
Private mvarEmployeeNo As String 'local copy
Private mvarDateOfBirth As Date 'local copy
Private mvarEngagementDate As Date 'local copy
Private mvarPOBox As String 'local copy
Private mvarMobilePhone As String 'local copy
'local variable(s) to hold property value(s)
Private mvarEmployeeSearch As String 'local copy
Public Property Let EmployeeSearch(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmployeeSearch = 5
    mvarEmployeeSearch = vData
End Property


Public Property Get EmployeeSearch() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmployeeSearch
    EmployeeSearch = mvarEmployeeSearch
End Property




Public Property Let MobilePhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MobilePhone = 5
    mvarMobilePhone = vData
End Property


Public Property Get MobilePhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MobilePhone
    MobilePhone = mvarMobilePhone
End Property



Public Property Let POBox(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.POBox = 5
    mvarPOBox = vData
End Property


Public Property Get POBox() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.POBox
    POBox = mvarPOBox
End Property



Public Property Let EngagementDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EngagementDate = 5
    mvarEngagementDate = vData
End Property


Public Property Get EngagementDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EngagementDate
    EngagementDate = mvarEngagementDate
End Property



Public Property Let DateOfBirth(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateOfBirth = 5
    mvarDateOfBirth = vData
End Property


Public Property Get DateOfBirth() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateOfBirth
    DateOfBirth = mvarDateOfBirth
End Property




Public Property Let EmployeeNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmployeeNo = 5
    mvarEmployeeNo = vData
End Property


Public Property Get EmployeeNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmployeeNo
    EmployeeNo = mvarEmployeeNo
End Property



Public Property Let WorkingStatus(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WorkingStatus = 5
    mvarWorkingStatus = vData
End Property


Public Property Get WorkingStatus() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WorkingStatus
    WorkingStatus = mvarWorkingStatus
End Property


Public Property Let Notes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Notes = 5
    mvarNotes = vData
End Property


Public Property Get Notes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Notes
    Notes = mvarNotes
End Property



Public Property Let Photo(ByVal vData As String)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax:  x.Photo = Form1
    mvarPhoto = vData
End Property


Public Property Get Photo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Photo
    Photo = mvarPhoto
End Property


Public Property Let HomePhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HomePhone = 5
    mvarHomePhone = vData
End Property


Public Property Get HomePhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HomePhone
    HomePhone = mvarHomePhone
End Property



Public Property Let Country(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Country = 5
    mvarCountry = vData
End Property


Public Property Get Country() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Country
    Country = mvarCountry
End Property



Public Property Let Address(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Address = 5
    mvarAddress = vData
End Property


Public Property Get Address() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Address
    Address = mvarAddress
End Property


Public Property Let TitleOfCourtesy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TitleOfCourtesy = 5
    mvarTitleOfCourtesy = vData
End Property


Public Property Get TitleOfCourtesy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TitleOfCourtesy
    TitleOfCourtesy = mvarTitleOfCourtesy
End Property




Public Property Let FirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstName = 5
    mvarFirstName = vData
End Property


Public Property Get FirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    FirstName = mvarFirstName
End Property



Public Property Let LastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastName = 5
    mvarLastName = vData
End Property


Public Property Get LastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastName
    LastName = mvarLastName
End Property



Public Property Let EmployeeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmployeeID = 5
    mvarEmployeeID = vData
End Property


Public Property Get EmployeeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmployeeID
    EmployeeID = mvarEmployeeID
End Property

Private Sub Class_Initialize()
'On Error GoTo errHandler

    Set con_Obj = mdl_MAIN.fn_OPEN_CONNECTION

EXITPROCEDURE:
    Exit Sub
    
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_Functions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub Class_Terminate()
On Error GoTo errHandler

    Call mdl_MAIN.fn_CLOSE_CONNECTION(con_Obj)

EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub
Public Function fn_LOAD_EMPLOYEES(Optional lngEmployeeID As Long) As ADODB.Recordset

    If lngEmployeeID = 0 Then
        db.strQueryString = "SELECT * FROM vw_EMPLOYEES "
        Else
            db.strQueryString = "SELECT * FROM vw_EMPLOYEES WHERE EmployeeID = " & lngEmployeeID
    End If

    Set fn_LOAD_EMPLOYEES = db.fn_MANAGE_RECORD

End Function

Public Function fn_LOAD_PICTURE(Optional lngEmployeeID As Long) As ADODB.Recordset

    If lngEmployeeID <> 0 Then
        db.strQueryString = "SELECT Photo FROM tb_EMPLOYEES WHERE EmployeeID = " & lngEmployeeID
    End If

    Set fn_LOAD_PICTURE = db.fn_MANAGE_RECORD

End Function

Public Function fn_SEARCH_EMPLOYEES(Optional strSearch As String, Optional lbl As Label) As ADODB.Recordset

    If strSearch <> "" Then
        lbl.Caption = "SELECT * FROM vw_EMPLOYEES WHERE  " & strSearch
        db.strQueryString = lbl.Caption
        Set fn_SEARCH_EMPLOYEES = db.fn_MANAGE_RECORD
    End If

End Function

Public Function fn_LOAD_REPORT(Optional lngEmployeeID As Long, Optional strSearch As String) As ADODB.Recordset

    If lngEmployeeID = 0 Then
        db.strQueryString = "SELECT * FROM vw_EMPLOYEES WHERE  " & strSearch
        Else
            db.strQueryString = "SELECT * FROM vw_EMPLOYEES WHERE  " & strSearch & " AND EmployeeID=" & lngEmployeeID
    End If

    Set fn_LOAD_REPORT = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_EMPLOYEE_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_EMPLOYEES " _
    & "(EmployeeNo,FirstName,LastName,TitleOfCourtesy,DateOfBirth,EngagementDate,Address,POBox,Country,HomePhone,MobilePhone,Notes,WorkingStatus,Photo) " _
    & " VALUES " _
    & "('" & mvarEmployeeNo & "'," _
    & "'" & mvarFirstName & "'," _
    & "'" & mvarLastName & "'," _
    & mvarTitleOfCourtesy & "," _
    & "'" & mvarDateOfBirth & "'," _
    & "'" & mvarEngagementDate & "'," _
    & "'" & mvarAddress & "'," _
    & "'" & mvarPOBox & "'," _
    & "'" & mvarCountry & "'," _
    & "'" & mvarHomePhone & "'," _
    & "'" & mvarMobilePhone & "'," _
    & "'" & mvarNotes & "'," _
    & mvarWorkingStatus & "," _
    & "'" & mvarPhoto & "')"

    Set fn_SAVE_EMPLOYEE_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_EMPLOYEE_RECORDS(Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_EMPLOYEES SET " _
    & "FirstName = '" & mvarFirstName & "'," _
    & "LastName = '" & mvarLastName & "'," _
    & "TitleOfCourtesy = " & mvarTitleOfCourtesy & "," _
    & "DateOfBirth = '" & mvarDateOfBirth & "'," _
    & "EngagementDate = '" & mvarEngagementDate & "'," _
    & "Address = '" & mvarAddress & "'," _
    & "POBox = '" & mvarPOBox & "'," _
    & "Country = '" & mvarCountry & "'," _
    & "HomePhone = '" & mvarHomePhone & "'," _
    & "MobilePhone = '" & mvarMobilePhone & "'," _
    & "Notes = '" & mvarNotes & "'," _
    & "WorkingStatus = " & mvarWorkingStatus & "," _
    & "Photo = '" & mvarPhoto & "'" _
    & " WHERE EmployeeID = " & lngEmployeeID
    
    Set fn_UPDATE_EMPLOYEE_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_DELETE_EMPLOYEE_RECORDS(Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = " UPDATE tb_EMPLOYEES SET " _
    & "WorkingStatus = 1 " _
    & "Where EmployeeID = " & lngEmployeeID

    Set fn_DELETE_EMPLOYEE_RECORDS = db.fn_MANAGE_RECORD

End Function


Public Function fn_DELETE_EMPLOYEE_PICTURE(Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = " UPDATE tb_EMPLOYEES SET " _
    & "Photo =  '" & "" & "'" _
    & "Where EmployeeID = " & lngEmployeeID

    Set fn_DELETE_EMPLOYEE_PICTURE = db.fn_MANAGE_RECORD

End Function

Public Function fn_AUTOGEN() As String
 
    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(EmployeeNo) FROM tb_EMPLOYEES"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = "E00001"
        Else
            fn_AUTOGEN = "E" & Format(Right(Trim(rec(0)), 5) + 1, "00000")
        End If
    Else
        fn_AUTOGEN = "E00001"
    End If

End Function
