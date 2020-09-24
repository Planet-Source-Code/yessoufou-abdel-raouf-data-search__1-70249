VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_EMPLOYEES_SEARCH 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CUSTOMER SEARCH"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   12180
   Begin VB.Frame Frame1 
      BackColor       =   &H00F6F8F8&
      Height          =   8715
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   12195
      Begin VB.Frame Frame3 
         BackColor       =   &H00F6F8F8&
         Height          =   5745
         Left            =   60
         TabIndex        =   2
         Top             =   2850
         Width           =   12075
         Begin MSComctlLib.ListView lvw 
            Height          =   4995
            Left            =   90
            TabIndex        =   3
            Top             =   600
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   8811
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
            NumItems        =   14
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
               Text            =   "Ageing"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Home Phone"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Mobile Phone"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "POBox"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Country"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Address"
               Object.Width           =   2540
            EndProperty
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
            Left            =   120
            TabIndex        =   4
            Top             =   210
            Width           =   1605
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F6F8F8&
         Height          =   2625
         Left            =   60
         TabIndex        =   1
         Top             =   90
         Width           =   12075
         Begin VB.CommandButton Command1 
            Caption         =   "Add New Employee"
            Height          =   495
            Left            =   150
            TabIndex        =   31
            Top             =   1980
            Width           =   2055
         End
         Begin VB.CommandButton cmdSelected 
            Caption         =   "Preview Selected"
            Height          =   525
            Left            =   9720
            TabIndex        =   30
            Top             =   1950
            Width           =   2055
         End
         Begin VB.CommandButton cmdPreviewAll 
            Caption         =   "Preview All"
            Height          =   525
            Left            =   7326
            TabIndex        =   29
            Top             =   1950
            Width           =   2055
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   525
            Left            =   4934
            TabIndex        =   28
            Top             =   1950
            Width           =   2055
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   525
            Left            =   2520
            TabIndex        =   27
            Top             =   1950
            Width           =   2055
         End
         Begin VB.TextBox txtFirstName 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1020
            Width           =   4425
         End
         Begin VB.TextBox txtLastName 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1395
            Width           =   4425
         End
         Begin VB.ComboBox cboTitleOfCourtesy 
            Height          =   315
            Left            =   7380
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   4425
         End
         Begin VB.TextBox txtHomePhone 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   7380
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1035
            Width           =   4425
         End
         Begin VB.TextBox txtMobilePhone 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   7380
            MaxLength       =   20
            TabIndex        =   6
            Top             =   1410
            Width           =   4425
         End
         Begin VB.TextBox txtEmployeeNo 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1410
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   4425
         End
         Begin MSComCtl2.DTPicker DTDateOfBirthFrom 
            Height          =   315
            Left            =   7830
            TabIndex        =   11
            Top             =   630
            Width           =   1755
            _ExtentX        =   3096
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
            Top             =   630
            Width           =   1755
            _ExtentX        =   3096
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
            Left            =   4080
            TabIndex        =   13
            Top             =   630
            Width           =   1755
            _ExtentX        =   3096
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
            Left            =   10050
            TabIndex        =   14
            Top             =   630
            Width           =   1755
            _ExtentX        =   3096
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
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   180
            X2              =   11790
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   345
            Left            =   180
            TabIndex        =   26
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   345
            Left            =   180
            TabIndex        =   25
            Top             =   1395
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Title Of Courtesy"
            Height          =   345
            Left            =   6090
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Birth:"
            Height          =   195
            Left            =   6120
            TabIndex        =   23
            Top             =   630
            Width           =   960
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Phone"
            Height          =   345
            Left            =   6120
            TabIndex        =   22
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile Phone"
            Height          =   345
            Left            =   6120
            TabIndex        =   21
            Top             =   1410
            Width           =   1125
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Engaged:"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee No"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   3810
            TabIndex        =   18
            Top             =   630
            Width           =   195
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   9780
            TabIndex        =   17
            Top             =   630
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   1410
            TabIndex        =   16
            Top             =   630
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   7380
            TabIndex        =   15
            Top             =   630
            Width           =   345
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

Private Sub sub_FILL_LIST(strSearch As String)
    Dim rec As New ADODB.Recordset
    lvw.ListItems.Clear
    
    Set rec = cls_EMPLOYEES_Obj.fn_SEARCH_EMPLOYEES(strSearch)
    
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
            lstItem.ListSubItems.Add , , Day(Date) - Day(rec!EngagementDate)
            lstItem.ListSubItems.Add , , rec!HomePhone
            lstItem.ListSubItems.Add , , rec!MobilePhone
            lstItem.ListSubItems.Add , , rec!POBox
            lstItem.ListSubItems.Add , , rec!Country
            lstItem.ListSubItems.Add , , rec!Address
            rec.MoveNext
        Loop
        
        lblResult.Caption = lvw.ListItems.Count & " Record(s) Found"
        
    End If
    
    
    
End Sub

Private Sub cboTitleOfCourtesy_Click()
    If cboTitleOfCourtesy.ListIndex = -1 Then Exit Sub
    lngTitle = cboTitleOfCourtesy.ItemData(cboTitleOfCourtesy.ListIndex)
End Sub

Private Sub cmdClear_Click()
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    resetDatePicker
    lvw.ListItems.Clear
    lblResult.Caption = "No Record(s) Found"
End Sub

Private Sub cmdSearch_Click()
    
    
    Dim strWhere As String
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
            strWhere = strWhere & "EngagementDate >= '" & DTEngagementDateFrom.Value & "'"
            Else
                strWhere = strWhere & " AND EngagementDate >= '" & DTEngagementDateFrom.Value & "'"
        End If
    End If

    If IsNull(DTEngagementDateTo.Value) = False Then
        If strWhere = "" Then
            strWhere = strWhere & "EngagementDate <= '" & DTEngagementDateTo.Value & "'"
            Else
                strWhere = strWhere & " AND EngagementDate <= '" & DTEngagementDateTo.Value & "'"
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
    
    If Trim(lngTitle) > 0 Then
        If strWhere = "" Then
            strWhere = strWhere & "TitleID = " & lngTitle
            Else
                strWhere = strWhere & " and " & " TitleID = " & lngTitle
        End If
    End If
    
    If IsNull(DTDateOfBirthFrom.Value) = False Then
        If strWhere = "" Then
            strWhere = strWhere & "DateOfBirth >= '" & DTDateOfBirthFrom.Value & "'"
            Else
                strWhere = strWhere & " AND DateOfBirth >= '" & DTDateOfBirthFrom.Value & "'"
        End If
    End If

    If IsNull(DTDateOfBirthTo.Value) = False Then
        If strWhere = "" Then
            strWhere = strWhere & "DateOfBirth <= '" & DTDateOfBirthTo.Value & "'"
            Else
                strWhere = strWhere & " AND DateOfBirth <= '" & DTDateOfBirthTo.Value & "'"
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
    
    If strWhere <> "" Then
        sub_FILL_LIST (strWhere)
    End If
    
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
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

Private Sub txtEmployeeNo_Validate(Cancel As Boolean)
    txtEmployeeNo.Text = StrConv(txtEmployeeNo, vbUpperCase)
End Sub

