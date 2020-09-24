VERSION 5.00
Begin VB.MDIForm frm_MAIN 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      Begin DATA_SEARCH.b8Container ctn1 
         Height          =   615
         Left            =   -30
         TabIndex        =   2
         Top             =   -30
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   1085
         BackColor       =   16185592
         Begin DATA_SEARCH.lvButtons_H cmdEmployee 
            Height          =   435
            Left            =   120
            TabIndex        =   8
            Top             =   90
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   767
            Caption         =   "EMPLOYEES"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_MAIN.frx":0000
         End
         Begin DATA_SEARCH.lvButtons_H cmdSearchEmployee 
            Height          =   435
            Left            =   3060
            TabIndex        =   9
            Top             =   90
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   767
            Caption         =   "SEARCH EMPLOYEES"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_MAIN.frx":031A
         End
         Begin VB.Image imgTop 
            Height          =   495
            Left            =   60
            Picture         =   "frm_MAIN.frx":0634
            Stretch         =   -1  'True
            Top             =   60
            Width           =   15150
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   2475
      Width           =   4680
      Begin DATA_SEARCH.b8Container ctn2 
         Height          =   615
         Left            =   -30
         TabIndex        =   3
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1085
         BackColor       =   16185592
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date && Time:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   9600
            TabIndex        =   7
            Top             =   180
            Width           =   1605
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   9120
            Picture         =   "frm_MAIN.frx":28C3
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lblName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   6480
            TabIndex        =   6
            Top             =   180
            Width           =   2565
         End
         Begin VB.Image Image4 
            Height          =   660
            Left            =   4200
            Picture         =   "frm_MAIN.frx":2D05
            Stretch         =   -1  'True
            Top             =   0
            Width           =   750
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   0
            Picture         =   "frm_MAIN.frx":49CF
            Stretch         =   -1  'True
            Top             =   60
            Width           =   690
         End
         Begin VB.Label lblRole 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   855
            TabIndex        =   5
            Top             =   210
            Width           =   2985
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   4950
            TabIndex        =   4
            Top             =   180
            Width           =   1455
         End
         Begin VB.Image imgBotom 
            Height          =   495
            Left            =   0
            Picture         =   "frm_MAIN.frx":A5E1
            Stretch         =   -1  'True
            Top             =   60
            Width           =   15150
         End
      End
   End
End
Attribute VB_Name = "frm_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdEmployee_Click()
    With frm_EMPLOYEES
        .Show
        .SetFocus
    End With
End Sub

Private Sub cmdSearchEmployee_Click()
    With frm_EMPLOYEES_SEARCH
        .Show
        .SetFocus
    End With
End Sub

Private Sub MDIForm_Resize()
    ctn1.Width = frm_MAIN.Width
    imgTop.Width = ctn1.Width
    ctn2.Width = frm_MAIN.Width
    imgBotom.Width = ctn2.Width
End Sub
