VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_EMPLOYEES_SEARCH 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "CUSTOMER SEARCH"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00F6F8F8&
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   14985
      Begin DATA_SEARCH.lvButtons_H cmdClose 
         Height          =   345
         Left            =   14340
         TabIndex        =   33
         Top             =   0
         Width           =   555
         _extentx        =   979
         _extenty        =   609
         caption         =   "X"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frm_EMPLOYEES_SEARCH.frx":0000
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F6F8F8&
         Caption         =   "FILTERING"
         ForeColor       =   &H00FF0000&
         Height          =   2505
         Left            =   180
         TabIndex        =   1
         Top             =   150
         Width           =   14715
         Begin VB.TextBox txtFirstName 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1050
            Width           =   4965
         End
         Begin VB.TextBox txtLastName 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1425
            Width           =   4965
         End
         Begin VB.ComboBox cboTitleOfCourtesy 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   270
            Width           =   4965
         End
         Begin VB.TextBox txtHomePhone 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   8070
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1065
            Width           =   4965
         End
         Begin VB.TextBox txtMobilePhone 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   8070
            MaxLength       =   20
            TabIndex        =   6
            Top             =   1440
            Width           =   4965
         End
         Begin VB.TextBox txtEmployeeNo 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1410
            MaxLength       =   6
            TabIndex        =   5
            Top             =   270
            Width           =   4965
         End
         Begin MSComCtl2.DTPicker DTDateOfBirthFrom 
            Height          =   315
            Left            =   8490
            TabIndex        =   11
            Top             =   660
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   16449537
            CurrentDate     =   39095
         End
         Begin MSComCtl2.DTPicker DTEngagementDateFrom 
            Height          =   315
            Left            =   1860
            TabIndex        =   12
            Top             =   660
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   16449537
            CurrentDate     =   39095
         End
         Begin MSComCtl2.DTPicker DTEngagementDateTo 
            Height          =   315
            Left            =   4350
            TabIndex        =   13
            Top             =   660
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   16449537
            CurrentDate     =   39095
         End
         Begin MSComCtl2.DTPicker DTDateOfBirthTo 
            Height          =   315
            Left            =   11040
            TabIndex        =   14
            Top             =   660
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   16449537
            CurrentDate     =   39095
         End
         Begin DATA_SEARCH.lvButtons_H cmdSearch 
            Default         =   -1  'True
            Height          =   585
            Left            =   1410
            TabIndex        =   29
            Top             =   1830
            Width           =   2325
            _extentx        =   4101
            _extenty        =   1032
            caption         =   "Search"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frm_EMPLOYEES_SEARCH.frx":002C
            mode            =   0
            value           =   0   'False
            image           =   "frm_EMPLOYEES_SEARCH.frx":0058
            imgsize         =   32
            cback           =   -2147483633
            micon           =   "frm_EMPLOYEES_SEARCH.frx":0CAA
            mpointer        =   99
         End
         Begin DATA_SEARCH.lvButtons_H cmdClear 
            Height          =   585
            Left            =   4020
            TabIndex        =   30
            Top             =   1830
            Width           =   2325
            _extentx        =   4101
            _extenty        =   1032
            caption         =   "Clear"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frm_EMPLOYEES_SEARCH.frx":0FC4
            mode            =   0
            value           =   0   'False
            image           =   "frm_EMPLOYEES_SEARCH.frx":0FF0
            imgsize         =   32
            cback           =   -2147483633
            micon           =   "frm_EMPLOYEES_SEARCH.frx":1E42
            mpointer        =   99
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   345
            Left            =   180
            TabIndex        =   26
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   345
            Left            =   180
            TabIndex        =   25
            Top             =   1425
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Title Of Courtesy"
            Height          =   345
            Left            =   6750
            TabIndex        =   24
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Birth:"
            Height          =   195
            Left            =   6780
            TabIndex        =   23
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Phone"
            Height          =   345
            Left            =   6780
            TabIndex        =   22
            Top             =   1065
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile Phone"
            Height          =   345
            Left            =   6780
            TabIndex        =   21
            Top             =   1440
            Width           =   1125
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Engaged:"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   660
            Width           =   1080
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee No"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   4080
            TabIndex        =   18
            Top             =   660
            Width           =   195
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   10740
            TabIndex        =   17
            Top             =   660
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   1410
            TabIndex        =   16
            Top             =   660
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   8040
            TabIndex        =   15
            Top             =   660
            Width           =   345
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F6F8F8&
         Caption         =   "SELECT STATEMENT"
         ForeColor       =   &H00FF0000&
         Height          =   1035
         Left            =   180
         TabIndex        =   27
         Top             =   2700
         Width           =   14715
         Begin VB.Label lblSearch 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   1  'Fixed Single
            Height          =   645
            Left            =   150
            TabIndex        =   28
            Top             =   270
            Width           =   14385
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F6F8F8&
         Caption         =   "RESULTS"
         ForeColor       =   &H00FF0000&
         Height          =   5115
         Left            =   180
         TabIndex        =   2
         Top             =   3750
         Width           =   14685
         Begin MSComctlLib.ListView lvw 
            Height          =   3705
            Left            =   120
            TabIndex        =   3
            Top             =   540
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   6535
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EmployeeID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Employee No"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Title"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "First Name"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Last Name"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Date Of Birth"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Age"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Date Engaged"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Home Phone"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Mobile Phone"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "POBox"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Country"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Address"
               Object.Width           =   2540
            EndProperty
         End
         Begin DATA_SEARCH.lvButtons_H cmdPreviewAll 
            Height          =   585
            Left            =   1530
            TabIndex        =   31
            Top             =   4350
            Width           =   2325
            _extentx        =   4101
            _extenty        =   1032
            caption         =   "Preview All"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frm_EMPLOYEES_SEARCH.frx":215C
            mode            =   0
            value           =   0   'False
            image           =   "frm_EMPLOYEES_SEARCH.frx":2188
            imgsize         =   32
            cback           =   -2147483633
            micon           =   "frm_EMPLOYEES_SEARCH.frx":2652
            mpointer        =   99
         End
         Begin DATA_SEARCH.lvButtons_H cmdSelected 
            Height          =   585
            Left            =   4110
            TabIndex        =   32
            Top             =   4350
            Width           =   2325
            _extentx        =   4101
            _extenty        =   1032
            caption         =   "Preview Selected"
            capalign        =   2
            backstyle       =   2
            cgradient       =   0
            font            =   "frm_EMPLOYEES_SEARCH.frx":296C
            mode            =   0
            value           =   0   'False
            image           =   "frm_EMPLOYEES_SEARCH.frx":2998
            imgsize         =   32
            cback           =   -2147483633
            micon           =   "frm_EMPLOYEES_SEARCH.frx":2E62
            mpointer        =   99
         End
         Begin VB.Image imgPicture 
            Height          =   3675
            Left            =   11820
            Picture         =   "frm_EMPLOYEES_SEARCH.frx":317C
            Stretch         =   -1  'True
            Top             =   540
            Width           =   2685
         End
         Begin VB.Label lblResult 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records Found"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   4
            Top             =   210
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frm_EMPLOYEES_SEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lngTitle As Long
Dim lngSelectedEmployeeID
Dim strWhere As String

Private Sub sub_FILL_LIST(strSearch As String)
    Dim rec As New ADODB.Recordset
    lvw.ListItems.Clear
    
    Set rec = cls_EMPLOYEES_Obj.fn_SEARCH_EMPLOYEES(strSearch, lblSearch)
    
    lblResult.Caption = "No Record(s) Found"
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            Set lstItem = lvw.ListItems.Add(, , rec!EmployeeID)
            lstItem.ListSubItems.Add , , rec!EmployeeNo & ""
            lstItem.ListSubItems.Add , , rec!TitleName & ""
            lstItem.ListSubItems.Add , , rec!FirstName & ""
            lstItem.ListSubItems.Add , , rec!LastName & ""
            lstItem.ListSubItems.Add , , rec!DateOfBirth & ""
            lstItem.ListSubItems.Add , , Year(Date) - Year(rec!DateOfBirth)
            lstItem.ListSubItems.Add , , rec!EngagementDate & ""
            'lstItem.ListSubItems.Add , , Day(Date) - Day(rec!EngagementDate)
            lstItem.ListSubItems.Add , , rec!HomePhone & ""
            lstItem.ListSubItems.Add , , rec!MobilePhone & ""
            lstItem.ListSubItems.Add , , rec!POBox & ""
            lstItem.ListSubItems.Add , , rec!Country & ""
            lstItem.ListSubItems.Add , , rec!Address & ""
            rec.MoveNext
        Loop
        
        lblResult.Caption = lvw.ListItems.Count & " Record(s) Found"
        cmdPreviewAll.Enabled = True
        cmdSelected.Enabled = True
        
    End If
    
    
    
End Sub

Private Sub cboTitleOfCourtesy_Click()
    If cboTitleOfCourtesy.ListIndex = -1 Then Exit Sub
    lngTitle = cboTitleOfCourtesy.ItemData(cboTitleOfCourtesy.ListIndex)
End Sub



Private Sub cmdClear_Click()
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    lngTitle = 0
    resetDatePicker
    lvw.ListItems.Clear
    lblSearch.Caption = ""
    lblResult.Caption = "No Record(s) Found"
    cmdPreviewAll.Enabled = False
    cmdSelected.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreviewAll_Click()
    
    If lvw.ListItems.Count <= 0 Then
        MsgBox "There is no data to display. Please retry"
        Exit Sub
    End If
    showReport (0)

End Sub

Private Sub showReport(lngSelectedEmployeeID As Long)
On Error GoTo errHandler

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_REPORT(lngSelectedEmployeeID, strWhere)
    
    
    With rpt_EMPLOYEES
        Set .DataSource = rec
        If rec!Photo = "" Or IsNull(rec!Photo) Then
            Set .Sections("section1").Controls.Item("pic").Picture = LoadPicture(strPicturePath & "anonymous.jpg")
        Else
            Set .Sections("section1").Controls.Item("pic").Picture = LoadPicture(strPicturePath & rec!Photo)
        End If
        .Show
    End With
    
    
EXITPROCEDURE:
    Set rpt_EMPLOYEES = Nothing
    Exit Sub
    
errHandler:
    If Err.Number = 53 Then
        Set rpt_EMPLOYEES.Sections("section1").Controls.Item("pic").Picture = LoadPicture(strPicturePath & "anonymous.jpg")
        Exit Sub
    End If
    MsgBox "Error Occurred while loading picture", vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdSearch_Click()
    
    
    
    strWhere = ""
    
    If Trim(txtEmployeeNo.Text) <> "" Then
        If strWhere = "" Then
            strWhere = strWhere & "EmployeeNO = '" & Trim(txtEmployeeNo.Text) & "'"
            Else
                strWhere = strWhere & " and " & " EmployeeNO = '" & Trim(txtEmployeeNo.Text) & "'"
        End If
    End If
    
    If IsNull(DTEngagementDateFrom.Value) = False Then
        If strWhere = "" Then
            strWhere = strWhere & "EngagementDate >= #" & DTEngagementDateFrom.Value & "#"
            Else
                strWhere = strWhere & " AND EngagementDate >= #" & DTEngagementDateFrom.Value & "#"
        End If
    End If

    If IsNull(DTEngagementDateTo.Value) = False Then
        If strWhere = "" Then
            strWhere = strWhere & "EngagementDate <= #" & DTEngagementDateTo.Value & "#"
            Else
                strWhere = strWhere & " AND EngagementDate <= #" & DTEngagementDateTo.Value & "#"
        End If
    End If
    
    If Trim(txtFirstName.Text) <> "" Then
        If strWhere = "" Then
            strWhere = strWhere & "FirstName = '" & Trim(txtFirstName.Text) & "'"
            Else
                strWhere = strWhere & " and " & " FirstName = '" & Trim(txtFirstName.Text) & "'"
        End If
    End If
    
    If Trim(txtLastName.Text) <> "" Then
        If strWhere = "" Then
            strWhere = strWhere & "LastName = '" & Trim(txtLastName.Text) & "'"
            Else
                strWhere = strWhere & " and " & " LastName = '" & Trim(txtLastName.Text) & "'"
        End If
    End If
    
    If lngTitle > 0 Then
        If strWhere = "" Then
            strWhere = strWhere & "TitleID = " & lngTitle
            Else
                strWhere = strWhere & " and " & " TitleID = " & lngTitle
        End If
    End If
    
    If IsNull(DTDateOfBirthFrom.Value) = False Then
        If strWhere = "" Then
            strWhere = strWhere & "DateOfBirth >= #" & DTDateOfBirthFrom.Value & "#"
            Else
                strWhere = strWhere & " AND DateOfBirth >= #" & DTDateOfBirthFrom.Value & "#"
        End If
    End If

    If IsNull(DTDateOfBirthTo.Value) = False Then
        If strWhere = "" Then
            strWhere = strWhere & "DateOfBirth <= #" & DTDateOfBirthTo.Value & "#"
            Else
                strWhere = strWhere & " AND DateOfBirth <= #" & DTDateOfBirthTo.Value & "#"
        End If
    End If
    
    If Trim(txtHomePhone.Text) <> "" Then
        If strWhere = "" Then
            strWhere = strWhere & "HomePhone = '" & Trim(txtHomePhone.Text) & "'"
            Else
                strWhere = strWhere & " and " & " HomePhone = '" & Trim(txtHomePhone.Text) & "'"
        End If
    End If
    
    If Trim(txtMobilePhone.Text) <> "" Then
        If strWhere = "" Then
            strWhere = strWhere & "MobilePhone = '" & Trim(txtMobilePhone.Text) & "'"
            Else
                strWhere = strWhere & " and " & " MobilePhone = '" & Trim(txtMobilePhone.Text) & "'"
        End If
    End If
    
    lvw.ListItems.Clear
    
    cmdPreviewAll.Enabled = False
    cmdSelected.Enabled = False
    
    If strWhere <> "" Then
        sub_FILL_LIST (strWhere)
    End If
    
End Sub



Private Sub cmdSelected_Click()


    If lngSelectedEmployeeID = 0 Then
        MsgBox "Please select the employee who's info you want to preview"
        Exit Sub
    End If
    
    showReport (lngSelectedEmployeeID)
    
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, 0
    Call sub_LOAD_TITLES(cboTitleOfCourtesy)
    resetDatePicker
End Sub


Private Sub resetDatePicker()
    DTEngagementDateFrom.Value = Date
    DTEngagementDateFrom.Value = Null
    DTEngagementDateTo.Value = Date
    DTEngagementDateTo.Value = Null
    DTDateOfBirthFrom.Value = Date
    DTDateOfBirthFrom.Value = Null
    DTDateOfBirthTo.Value = Date
    DTDateOfBirthTo.Value = Null
End Sub

Private Sub sub_LOAD_TITLES(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOADT_TITLES
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!TitleName
        cbo.ItemData(cbo.NewIndex) = rec!TitleID
        rec.MoveNext
    Loop

End Sub



Private Sub lvw_Click()
    If lvw.ListItems.Count <= 0 Then Exit Sub
    lngSelectedEmployeeID = lvw.SelectedItem.Text
    loadEmployeePicture (lngSelectedEmployeeID)

End Sub

Public Sub loadEmployeePicture(lngEmployeeID As Long)
On Error GoTo errHandler
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_PICTURE(lngEmployeeID)
    If rec!Photo = "" Or IsNull(rec!Photo) Then
        imgPicture.Picture = LoadPicture(strPicturePath & "anonymous.jpg")
        Else
            imgPicture.Picture = LoadPicture(strPicturePath & rec!Photo)
    End If

EXITPROCEDURE:
    Exit Sub

errHandler:
    If Err.Number = 53 Then
        imgPicture.Picture = LoadPicture(strPicturePath & "anonymous.jpg")
        Exit Sub
    End If
    MsgBox "Error Occurred while loading picture", vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    
    GoTo EXITPROCEDURE
End Sub
Private Sub txtEmployeeNo_Validate(Cancel As Boolean)
    txtEmployeeNo.Text = StrConv(txtEmployeeNo, vbUpperCase)
End Sub

