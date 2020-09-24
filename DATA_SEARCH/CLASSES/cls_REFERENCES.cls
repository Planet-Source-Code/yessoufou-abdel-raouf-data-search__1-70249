VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_REFERENCES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTitleID As Integer 'local copy
Private mvarTitleName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarPackageName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarExpenditureName As String 'local copy
Public Property Let ExpenditureName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureName = 5
    mvarExpenditureName = vData
End Property


Public Property Get ExpenditureName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureName
    ExpenditureName = mvarExpenditureName
End Property



Public Property Let PackageName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PackageName = 5
    mvarPackageName = vData
End Property


Public Property Get PackageName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PackageName
    PackageName = mvarPackageName
End Property



Public Property Let TitleName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReferenceValue = 5
    mvarTitleName = vData
End Property


Public Property Get TitleName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReferenceValue
    TitleName = mvarTitleName
End Property



Public Property Let TitleID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RefrenceID = 5
    mvarTitleID = vData
End Property


Public Property Get TitleID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RefrenceID
    TitleID = mvarTitleID
End Property


Public Function fn_LOADT_TITLES(Optional lngID As Long) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_TITLES ORDER BY TitleName ASC"
        Else
            db.strQueryString = "SELECT * FROM tb_TITLES WHERE TitleID = " & lngID & " ORDER BY TitleName ASC"
    End If
    
    Set fn_LOADT_TITLES = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_LOAD_PACKAGES(Optional lngID As Long) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_PACKAGES ORDER BY PackageName ASC"
        Else
            db.strQueryString = "SELECT * FROM tb_PACKAGES WHERE PackageID = " & lngID & " ORDER BY PackageName ASC"
    End If
    
    Set fn_LOAD_PACKAGES = db.fn_MANAGE_RECORD
    
End Function


Public Function fn_LOAD_EXPENDITURE_TYPE(Optional lngID As Long) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_EXPENDITURE_TYPE ORDER BY ExpenditureName ASC"
        Else
            db.strQueryString = "SELECT * FROM tb_EXPENDITURE_TYPE WHERE ExpenditureTypeID = " & lngID & " ORDER BY ExpenditureName ASC"
    End If
    
    Set fn_LOAD_EXPENDITURE_TYPE = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_SAVE_PACKAGE() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_PACKAGES " _
    & "(PackageName) " _
    & " VALUES " _
    & "('" & mvarPackageName & "')"

    Set fn_SAVE_PACKAGE = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_PACKAGE(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_PACKAGES SET " _
    & "PackageName = '" & mvarPackageName & "'" _
    & " WHERE PackageID = " & lngID
    
    Set fn_UPDATE_PACKAGE = db.fn_MANAGE_RECORD

End Function


Public Function fn_SAVE_EXPENDITURE() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_EXPENDITURE_TYPE " _
    & "(ExpenditureName) " _
    & " VALUES " _
    & "('" & mvarExpenditureName & "')"

    Set fn_SAVE_EXPENDITURE = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_EXPENDITURE(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_EXPENDITURE_TYPE SET " _
    & "ExpenditureName = '" & mvarExpenditureName & "'" _
    & " WHERE ExpenditureTypeID = " & lngID
    
    Set fn_UPDATE_EXPENDITURE = db.fn_MANAGE_RECORD

End Function


Public Function fn_LOAD_LEAVES(Optional lngLeaveID As Long) As ADODB.Recordset

    If lngLeaveID = 0 Then
        db.strQueryString = "SELECT * FROM tb_LEAVES ORDER BY LeaveName ASC"
        Else
            db.strQueryString = "SELECT * FROM tb_LEAVES WHERE LeaveID = " & lngLeaveID & " ORDER BY LeaveName ASC"
    End If

    Set fn_LOAD_LEAVES = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_LOAD_SUPERVISORS(Optional lngSupervisorID As Long) As ADODB.Recordset

    If lngSupervisorID = 0 Then
        db.strQueryString = "SELECT * FROM tb_EMPLOYEES ORDER BY FirstName ASC"
        Else
            db.strQueryString = "SELECT * FROM tb_EMPLOYEES WHERE EmployeeID = " & lngSupervisorID & " ORDER BY FirstName ASC"
    End If

    Set fn_LOAD_SUPERVISORS = db.fn_MANAGE_RECORD
    
End Function

