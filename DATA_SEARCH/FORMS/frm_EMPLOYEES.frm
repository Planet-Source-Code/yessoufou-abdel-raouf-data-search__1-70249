VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_EMPLOYEES 
   BackColor       =   &H00F6F8F8&
   BorderStyle     =   0  'None
   Caption         =   "Employees"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00F6F8F8&
      Height          =   8925
      Left            =   30
      TabIndex        =   10
      Top             =   90
      Width           =   14835
      Begin VB.Frame Frame4 
         BackColor       =   &H00F6F8F8&
         Height          =   885
         Left            =   450
         TabIndex        =   37
         Top             =   7890
         Width           =   13455
         Begin DATA_SEARCH.lvButtons_H cmdAddNew 
            Height          =   585
            Left            =   120
            TabIndex        =   39
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            Caption         =   "Add New"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":0000
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":0515
         End
         Begin DATA_SEARCH.lvButtons_H cmdEdit 
            Height          =   585
            Left            =   2388
            TabIndex        =   40
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            Caption         =   "Edit"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":082F
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":0D66
         End
         Begin DATA_SEARCH.lvButtons_H cmdSave 
            Height          =   585
            Left            =   4656
            TabIndex        =   41
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            Caption         =   "Save"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":1080
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":12E0
         End
         Begin DATA_SEARCH.lvButtons_H cmdCancel 
            Height          =   585
            Left            =   6924
            TabIndex        =   42
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            Caption         =   "Cancel"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":15FA
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":1B69
         End
         Begin DATA_SEARCH.lvButtons_H cmdSearch 
            Height          =   585
            Left            =   9192
            TabIndex        =   43
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            Caption         =   "Search"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":1E83
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":2AD5
         End
         Begin DATA_SEARCH.lvButtons_H cmdClose 
            Height          =   585
            Left            =   11460
            TabIndex        =   44
            Top             =   180
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":2DEF
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":337F
         End
      End
      Begin VB.Frame fraPicture 
         BackColor       =   &H00F6F8F8&
         Enabled         =   0   'False
         Height          =   7755
         Left            =   11160
         TabIndex        =   34
         Top             =   120
         Width           =   2745
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00F6F8F8&
            Height          =   2685
            Left            =   150
            ScaleHeight     =   2625
            ScaleWidth      =   2325
            TabIndex        =   35
            Top             =   210
            Width           =   2385
            Begin VB.Image imgPicture 
               Height          =   2655
               Left            =   0
               Picture         =   "frm_EMPLOYEES.frx":3699
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2325
            End
            Begin VB.Label lblAlerte 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No Picture"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   60
               TabIndex        =   36
               Top             =   1110
               Width           =   1965
            End
         End
         Begin MSComDlg.CommonDialog PictureDlg 
            Left            =   390
            Top             =   4830
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin DATA_SEARCH.lvButtons_H cmdLoadPicture 
            Height          =   495
            Left            =   150
            TabIndex        =   45
            Top             =   3570
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   873
            Caption         =   "Add/Change"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":5D51
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":7A5B
         End
         Begin DATA_SEARCH.lvButtons_H cmdRemovePicture 
            Height          =   495
            Left            =   150
            TabIndex        =   46
            Top             =   4110
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   873
            Caption         =   "Remove"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frm_EMPLOYEES.frx":7D75
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":8185
         End
         Begin VB.Label lblIdentification 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   150
            TabIndex        =   38
            Top             =   2880
            Width           =   2355
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F6F8F8&
         Height          =   7755
         Left            =   450
         TabIndex        =   32
         Top             =   120
         Width           =   3735
         Begin VB.ListBox lstEmployees 
            Height          =   7275
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame fraEmployeeDetails 
         BackColor       =   &H00F6F8F8&
         Enabled         =   0   'False
         Height          =   7755
         Left            =   4290
         TabIndex        =   13
         Top             =   120
         Width           =   6735
         Begin VB.Frame Frame2 
            BackColor       =   &H00F6F8F8&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7785
            Left            =   0
            TabIndex        =   14
            Top             =   -30
            Width           =   6735
            Begin VB.TextBox txtFirstName 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   2
               Top             =   1170
               Width           =   5025
            End
            Begin VB.TextBox txtLastName 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   3
               Top             =   1575
               Width           =   5025
            End
            Begin VB.ComboBox cboTitleOfCourtesy 
               Height          =   315
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1980
               Width           =   5025
            End
            Begin VB.TextBox txtHomePhone 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   6
               Top             =   2805
               Width           =   5025
            End
            Begin VB.TextBox txtMobilePhone 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   7
               Top             =   3210
               Width           =   5025
            End
            Begin VB.TextBox txtCountry 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   11
               Top             =   5310
               Width           =   5025
            End
            Begin VB.TextBox txtAddress 
               Height          =   1185
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   9
               Top             =   4035
               Width           =   5025
            End
            Begin VB.TextBox txtPOBox 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   8
               Top             =   3630
               Width           =   5025
            End
            Begin VB.TextBox txtNotes 
               Height          =   1185
               IMEMode         =   3  'DISABLE
               Left            =   1500
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   5730
               Width           =   5025
            End
            Begin VB.CheckBox ChkAlerte 
               BackColor       =   &H000000C0&
               Caption         =   "This employee is no more working here"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   180
               TabIndex        =   15
               Top             =   6960
               Visible         =   0   'False
               Width           =   6315
            End
            Begin MSComCtl2.DTPicker DTDateOfBirth 
               Height          =   315
               Left            =   1500
               TabIndex        =   5
               Top             =   2400
               Width           =   2505
               _ExtentX        =   4419
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
               Format          =   64552961
               CurrentDate     =   39095
            End
            Begin MSComCtl2.DTPicker DTEngagementDate 
               Height          =   315
               Left            =   1500
               TabIndex        =   1
               Top             =   750
               Width           =   2505
               _ExtentX        =   4419
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
               Format          =   64552961
               CurrentDate     =   39095
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
               Height          =   345
               Left            =   150
               TabIndex        =   31
               Top             =   1170
               Width           =   1005
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               Height          =   345
               Left            =   150
               TabIndex        =   30
               Top             =   1575
               Width           =   855
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Title Of Courtesy"
               Height          =   345
               Left            =   150
               TabIndex        =   29
               Top             =   1980
               Width           =   1455
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Of Birth"
               Height          =   195
               Left            =   150
               TabIndex        =   28
               Top             =   2400
               Width           =   915
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Home Phone"
               Height          =   345
               Left            =   150
               TabIndex        =   27
               Top             =   2805
               Width           =   1095
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile Phone"
               Height          =   345
               Left            =   150
               TabIndex        =   26
               Top             =   3210
               Width           =   1125
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Country"
               Height          =   345
               Left            =   150
               TabIndex        =   25
               Top             =   5310
               Width           =   855
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   345
               Left            =   150
               TabIndex        =   24
               Top             =   4035
               Width           =   855
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "P.O.Box"
               Height          =   345
               Left            =   150
               TabIndex        =   23
               Top             =   3630
               Width           =   1125
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Notes"
               Height          =   345
               Left            =   150
               TabIndex        =   22
               Top             =   5730
               Width           =   855
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Engaged"
               Height          =   195
               Left            =   150
               TabIndex        =   21
               Top             =   750
               Width           =   1035
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Employee No"
               Height          =   195
               Left            =   150
               TabIndex        =   20
               Top             =   360
               Width           =   945
            End
            Begin VB.Label lblEmployeeNo 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   1500
               TabIndex        =   0
               Top             =   360
               Width           =   5025
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
               Height          =   345
               Left            =   4260
               TabIndex        =   19
               Top             =   2400
               Width           =   915
            End
            Begin VB.Label lblAge 
               Alignment       =   2  'Center
               BackColor       =   &H00F6F8F8&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4800
               TabIndex        =   18
               Top             =   2400
               Width           =   1725
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Ageing"
               Height          =   345
               Left            =   4260
               TabIndex        =   17
               Top             =   750
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lblAgeing 
               Alignment       =   2  'Center
               BackColor       =   &H00F6F8F8&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4800
               TabIndex        =   16
               Top             =   750
               Visible         =   0   'False
               Width           =   1725
            End
         End
      End
   End
End
Attribute VB_Name = "frm_EMPLOYEES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngID As Long

Dim strPictureName As String

Dim blnEmployeeModify As Boolean
Dim blnEmployeeAdd As Boolean



Private Sub ChkAlerte_Click()
    If ChkAlerte.Value = 1 Then
        ChkAlerte.Caption = "This employee is no more working here"
        ChkAlerte.BackColor = &HC0&
        Else
            ChkAlerte.Caption = "This employee should continue working?"
            ChkAlerte.BackColor = &HC000&
    End If
End Sub

Private Sub chkWorkingStatus_Click()
    If chkWorkingStatus.Value = 0 Then
        chkWorkingStatus.Caption = "Employee still working"
        chkWorkingStatus.BackColor = vbGreen
        Else
            chkWorkingStatus.Caption = "This employee has been deleted"
            chkWorkingStatus.BackColor = vbRed
    End If
End Sub

Private Sub cmdCancel_Click()
    Call fn_DISABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    fraEmployeeDetails.Enabled = False
    fraPicture.Enabled = False
    lstEmployees.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
'On Error GoTo errHandler

    If txtFirstName.Text = "" Then
        MsgBox "Kindly choose the Employee to be Deleted.", vbExclamation, Title
        GoTo EXITPROCEDURE
    End If
    
    If MsgBox("Are you sure you want to delete  " & Trim(txtFirstName.Text) & " ?", vbQuestion + vbYesNo, Title) = vbNo Then
        
        GoTo EXITPROCEDURE
        
        Else
        
            Call cls_EMPLOYEES_Obj.fn_DELETE_EMPLOYEE_RECORDS(lngID)
            MsgBox "Employee successfully deleted.", vbExclamation, Title
            
    End If
    
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call sub_LOAD_TITLES(cboTitleOfCourtesy)
    Call sub_LOAD_EMPLOYEES(lstEmployees)
    
EXITPROCEDURE:
    Exit Sub
    
'errHandler:
'    MsgBox "An error occured.", vbCritical, title
'    Call MdlFunctions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, Me.Name, "cmdCategorySupprimer")
'    GoTo EXITPROCEDURE
End Sub

Private Sub cmdaddNew_Click()
    blnEmployeeAdd = True
    blnEmployeeModify = False
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call fn_ENABLE_CONTROLS
    lblIdentification.Caption = cls_EMPLOYEES_Obj.fn_AUTOGEN
    lblEmployeeNo.Caption = lblIdentification.Caption
    Call fn_UNSET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = True
    fraPicture.Enabled = True
    lstEmployees.Enabled = False
    DTEngagementDate.SetFocus
    lblAge.Caption = ""
End Sub

Private Sub cmdLoadPicture_Click()
On Error GoTo errHandler
    
'    imgPicture.Picture = LoadPicture()
    PictureDlg.ShowOpen
    If PictureDlg.FileName = "" Then GoTo EXITPROCEDURE
    imgPicture.Picture = LoadPicture(PictureDlg.FileName)

EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox "La photo n'est pas valide", vbCritical, "Connection"
    Call MdlFunctions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdRemovePicture_Click()
    If MsgBox("Are you sure you want to remove  " & Trim(txtFirstName.Text) & "'s picture ?", vbQuestion + vbYesNo, Title) = vbNo Then
        
        GoTo EXITPROCEDURE
        
        Else
        
            Call cls_EMPLOYEES_Obj.fn_DELETE_EMPLOYEE_PICTURE(lngID)
            strPictureName = Trim(cmdIdentification.Caption) & ".bmp"
            Kill strPicturePath & strPictureName
            imgPicture.Picture = LoadPicture()
            MsgBox "Employee picture successfully deleted.", vbExclamation, Title
            
            Call sub_LOAD_TITLES(cboTitleOfCourtesy)
        
            Call fn_DISABLE_CONTROLS
            Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
            fraEmployeeDetails.Enabled = False
            fraPicture.Enabled = False
            Call sub_LOAD_EMPLOYEES(lstEmployees)
            lstEmployees.Enabled = True
    End If
    
    
    
        
EXITPROCEDURE:
    Exit Sub
End Sub



Private Sub cmdSave_Click()

    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtFirstName, "Kindly enter the employee first name.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtLastName, "Kindly enter the employee last name.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboTitleOfCourtesy, "Kindly enter the employee title of courtesy.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_DATE_OF_BIRTH(DTDateOfBirth, "The employee Birth Date can not be today or Tomorrow.") Then GoTo EXITPROCEDURE

    

    With cls_EMPLOYEES_Obj
        .EmployeeNo = Trim(lblIdentification.Caption)
        .FirstName = Trim(txtFirstName.Text)
        .LastName = Trim(txtLastName.Text)
        .TitleOfCourtesy = cboTitleOfCourtesy.ItemData(cboTitleOfCourtesy.ListIndex)
        .DateOfBirth = DTDateOfBirth.Value
        .EngagementDate = DTEngagementDate.Value
        .Address = Trim(txtAddress.Text)
        .POBox = Trim(txtPOBox.Text)
        .Country = Trim(txtCountry.Text)
        .HomePhone = Trim(txtHomePhone.Text)
        .MobilePhone = Trim(txtMobilePhone.Text)
        .Notes = Trim(txtNotes.Text)

        If ChkAlerte.Value = 0 Then
            .WorkingStatus = 0
            Else
                .WorkingStatus = 1
        End If

        If imgPicture.Picture Then
            strPictureName = Trim(lblIdentification.Caption) & ".bmp"
            Call SavePicture(imgPicture, strPicturePath & strPictureName)
            .Photo = strPictureName
            Else
                .Photo = ""
        End If
    End With


    If blnEmployeeAdd = True And blnEmployeeModify = False Then
        cls_EMPLOYEES_Obj.fn_SAVE_EMPLOYEE_RECORDS
        MsgBox "Employee successfully saved.", vbExclamation, Title
        Else
            cls_EMPLOYEES_Obj.fn_UPDATE_EMPLOYEE_RECORDS (lngID)
            MsgBox "Employee successfully updated.", vbExclamation, Title
    End If


    Call fn_DISABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = False
    fraPicture.Enabled = False
    Call sub_LOAD_EMPLOYEES(lstEmployees)
    lstEmployees.Enabled = True
    
    
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    blnEmployeeAdd = False
    blnEmployeeModify = True
    Call fn_ENABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = True
    fraPicture.Enabled = True
    lstEmployees.Enabled = False
    txtFirstName.SetFocus
End Sub

Private Sub cmdSearch_Click()
    With frm_EMPLOYEES_SEARCH
        .Show
        .SetFocus
    End With
End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) / 2, 0
    Call fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call fn_DISABLE_CONTROLS
    Call sub_LOAD_TITLES(cboTitleOfCourtesy)
    Call sub_LOAD_EMPLOYEES(lstEmployees)


End Sub

Private Sub sub_LOAD_EMPLOYEES(Optional lst As ListBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_EMPLOYEES
    
    lst.Clear
    
    Do While Not rec.EOF
        lst.AddItem rec!TitleName & "  " & rec!FirstName & "  " & rec!LastName
        lst.ItemData(lst.NewIndex) = rec!EmployeeID
        rec.MoveNext
    Loop

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


Private Sub sub_LOAD_SUPERVISORS(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_SUPERVISORS(0)
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!FirstName & " " & rec!LastName
        cbo.ItemData(cbo.NewIndex) = rec!EmployeeID
        rec.MoveNext
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnEmployeeAdd = False
    blnEmployeeModify = False
End Sub

Private Sub lstEmployees_Click()
'    Call MdlFunctions.fnEmptyFields(Me)
    lngID = lstEmployees.ItemData(lstEmployees.ListIndex)
    Call subLoadEmployeeDetails(lngID)
End Sub

Private Sub subLoadEmployeeDetails(lngEmployeeID As Long)
On Error GoTo errHandler
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_EMPLOYEES(lngEmployeeID)
    
    lblIdentification.Caption = rec!EmployeeNo
    lblEmployeeNo.Caption = rec!EmployeeNo
    txtFirstName.Text = rec!FirstName & " "
    txtLastName.Text = rec!LastName & " "
    cboTitleOfCourtesy.ListIndex = fn_GET_LIST_INDEX(cboTitleOfCourtesy, rec!TitleOfCourtesy)
    DTDateOfBirth.Value = rec!DateOfBirth & " "
    lblAge.Caption = Year(Date) - Year(rec!DateOfBirth)
    DTEngagementDate.Value = rec!EngagementDate & " "
    'lblAgeing.Caption = Day(Date) - Day(rec!EngagementDate)
    txtAddress.Text = rec!Address & " "
    txtPOBox.Text = rec!POBox & " "
    txtCountry.Text = rec!Country & " "
    txtHomePhone.Text = rec!HomePhone & " "
    txtMobilePhone.Text = rec!MobilePhone & " "
    txtNotes.Text = rec!Notes & " "
    
    If rec!WorkingStatus = 1 Then
        ChkAlerte.Visible = True
        ChkAlerte.Value = 1
        Else
            ChkAlerte.Visible = False
    End If
    
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

Private Sub txtCountry_Validate(Cancel As Boolean)
    txtCountry.Text = StrConv(txtCountry, vbProperCase)
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtFirstName_Validate(Cancel As Boolean)
    txtFirstName.Text = StrConv(txtFirstName, vbProperCase)
End Sub

Private Sub txtHomePhone_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub



Public Function fn_DISABLE_CONTROLS()

    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    cmdEdit.Enabled = True
    cmdSearch.Enabled = True
    cmdAddNew.Enabled = True

    cmdLoadPicture.Enabled = False
    cmdRemovePicture.Enabled = False

End Function

'**********Function to enable some buttons***********
Public Function fn_ENABLE_CONTROLS()

    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    cmdEdit.Enabled = False
    cmdSearch.Enabled = False
    cmdAddNew.Enabled = False

    cmdLoadPicture.Enabled = True
    cmdRemovePicture.Enabled = True


End Function

Private Sub txtLastName_Validate(Cancel As Boolean)
    txtLastName.Text = StrConv(txtLastName, vbProperCase)
End Sub

Private Sub txtMobilePhone_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtPOBox_Validate(Cancel As Boolean)
    txtLastName.Text = StrConv(txtLastName, vbUpperCase)
End Sub

