VERSION 5.00
Begin VB.Form frm_PICTURE 
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   LinkTopic       =   "Form1"
   Picture         =   "frm_PICTURE.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00F6F8F8&
      Height          =   2685
      Left            =   0
      ScaleHeight     =   2625
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   0
      Width           =   2235
      Begin VB.Image imgPicture 
         Height          =   2685
         Left            =   -30
         Picture         =   "frm_PICTURE.frx":8CB0
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frm_PICTURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Form_Load()
    Move 600, 6000
End Sub
