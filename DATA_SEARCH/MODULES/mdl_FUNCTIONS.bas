Attribute VB_Name = "Mdl_FUNCTIONS"

'**********Function to highlight the content of a text***********
Public Function fn_HIGHLIGHT_TEXT(ByVal txt As TextBox)

    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
    txt.SetFocus

End Function

'**********Function to require data in a textbox***********
Public Function fn_REQUIRE_TEXT_FIELD(ByVal txt As TextBox, strMessage As String) As Boolean
    
    fn_REQUIRE_TEXT_FIELD = False
    
    If txt.Text = "" Then
        fn_REQUIRE_TEXT_FIELD = True
        MsgBox strMessage, vbExclamation, Title
        txt.SetFocus
    End If
    
End Function

Public Function fn_REQUIRE_COMBO_FIELD(ByVal cbo As ComboBox, strMessage As String) As Boolean
    fn_REQUIRE_COMBO_FIELD = False
    
    If cbo.ListIndex = -1 Then
        fn_REQUIRE_COMBO_FIELD = True
        MsgBox strMessage, vbExclamation, Title
        cbo.SetFocus
    End If
    
End Function

Public Function fn_FILL_ALL_FIEL(frm As Form) As Boolean
    
    fn_FILL_ALL_FIEL = False
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Text = "" Then
                fn_FILL_ALL_FIEL = True
            End If
        End If
    Next

End Function


Public Function fn_REQUIRE_DATE_OF_BIRTH(ByVal dt As DTPicker, strMessage As String) As Boolean
    fn_REQUIRE_DATE_OF_BIRTH = False
    
    If dt.Value = Date Or dt.Value > Now Then
        fn_REQUIRE_DATE_OF_BIRTH = True
        MsgBox strMessage, vbExclamation, Title
        dt.SetFocus
    End If
    
End Function

'**********Function to accept only numeric characters***********
Public Sub sub_NUMERIC_ONLY(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select

End Sub

'**********Function to accept only numeric characters***********
Public Sub sub_FORM_SIZE(frm As Form)

    frm.Width = 12180
    frm.Height = 9300

End Sub

'**********Function to accept only alphabet***********
Public Sub sub_ALPHABET_ONLY(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9
            KeyAscii = 0
    End Select

End Sub

'**********Function to center a form on the screen***********
Public Sub sub_CENTER_FORM(ByVal frm As Form)

    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 3

End Sub

'**********Function to center a form on the screen***********
Public Sub sub_SET_FORM_POSITION(ByVal frm As Form)

    frm.Left = 0
    frm.Top = 0

End Sub

'**********Function to set the tree font***********
Public Sub sub_SET_TREEVIEW_FONT()

    XNode.Bold = True
    XNode.ForeColor = &HC00000

End Sub

'**********Function to unload a form***********
Public Sub sub_UNLOAD_FORMS(ByVal frm As Form)

    If MsgBox("Are you sure you want to close this window ?", vbQuestion + vbYesNo, Title) = vbNo Then
    
        Exit Sub
        
        Else
                
                Unload frm
    
    End If

End Sub

'**********Function to write all the errors that occur to a file***********
Public Function fn_WRITE_ERROR_TO_FILE(dDate As Date, dTime As Date, strErrDescription As String, strErrNumber As String, strModule As String, strProcedure As String) As String

    Set txtStream = file.OpenTextFile(App.Path & "\Errors\ErrorText.txt", ForAppending, True)
    txtStream.WriteLine "***********************************************"
    txtStream.WriteLine "*Date              :" & dDate
    txtStream.WriteLine "*Time              :" & dTime
    txtStream.WriteLine "*Error Description :" & strErrDescription
    txtStream.WriteLine "*Error Number      :" & strErrNumber
    txtStream.WriteLine "*Module            :" & strModule
    txtStream.WriteLine "*Procedure         :" & strProcedure
    txtStream.WriteLine "***********************************************"
    txtStream.Close
    
End Function


'**********Function to set the list index***********
Public Function fn_GET_LIST_INDEX(cbo As ComboBox, lngItemID As Long) As Long

    Dim lngIndex As Long

    fn_GET_LIST_INDEX = -1
    
    For lngIndex = 0 To cbo.ListCount - 1
        If cbo.ItemData(lngIndex) = lngItemID Then
            fn_GET_LIST_INDEX = lngIndex
        End If
    Next lngIndex

End Function

'**********Function to empty fields*******************
Public Sub sub_EMPTY_FIELS(frm As Form)

    For Each ctl In frm
        If TypeOf ctl Is Label And Left(ctl.Name, 3) = "lab" Then ctl.Caption = ""
        If TypeOf ctl Is TextBox Then ctl.Text = ""
        If TypeOf ctl Is ComboBox Then ctl.ListIndex = -1
        If TypeOf ctl Is DTPicker Then ctl.Value = Date
        If TypeOf ctl Is Image Then ctl.Picture = LoadPicture("")
    Next

End Sub


Public Function fn_CHECK_EMPTY_COMBO(cbo As ComboBox, strMessage As String) As Boolean

    fn_CHECK_EMPTY_COMBO = False

    If cbo.ListIndex = -1 Then
        fn_CHECK_EMPTY_COMBO = True
        MsgBox strMessage, vbExclamation, Title
        cbo.SetFocus
    End If


End Function

Public Function fn_CHECK_EMPTY_TEXT_BOX(txt As TextBox, strMessage As String) As Boolean

    fn_CHECK_EMPTY_TEXT_BOX = False

    If txt.Text = "" Then
        fn_CHECK_EMPTY_TEXT_BOX = True
        MsgBox strMessage, vbExclamation, Title
        txt.SetFocus
    End If


End Function


Public Function fn_CHECK_DATE_OF_BIRTH(dtp As DTPicker, strMessage As String) As Boolean

    fn_CHECK_DATE_OF_BIRTH = False

    If dtp.Value > Date Then
        fn_CHECK_DATE_OF_BIRTH = True
        MsgBox strMessage, vbExclamation, Title
        dtp.SetFocus
    End If


End Function


Public Sub sub_CLOSE_ALL_OPENED_FORMS()
    
    Dim i As Integer
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).Name <> "frm_MAIN" Then
            Unload Forms(i)
        End If
    Next i
    
    
End Sub


Public Function fn_SET_CONTROL_COLOR(frm As Form, Optional str As String)
    Dim col As Variant
    col = &HE0E0E0
    Select Case str
        Case "All"
                For Each ctl In frm
                    If TypeOf ctl Is TextBox Then ctl.BackColor = col
                    If TypeOf ctl Is ComboBox Then ctl.BackColor = col
                Next
        Case "Category"
                For Each ctl In frm
                    If TypeOf ctl Is TextBox And Left(Trim(ctl.Name), 4) = "txtC" Then ctl.BackColor = col
                Next
        
        Case "Product"
                For Each ctl In frm
                    If TypeOf ctl Is TextBox And Left(Trim(ctl.Name), 3) = "txt" Then ctl.BackColor = col
                    If TypeOf ctl Is ComboBox Then ctl.BackColor = col
                Next
        
    End Select
    

End Function


Public Function fn_UNSET_CONTROL_COLOR(frm As Form, str As String)
    Dim col As Variant
    col = &H80000005
    
    Select Case str
        Case "All"
                For Each ctl In frm
                    If TypeOf ctl Is TextBox Then ctl.BackColor = col
                    If TypeOf ctl Is ComboBox Then ctl.BackColor = col
                Next
        Case "Category"
                For Each ctl In frm
                    If TypeOf ctl Is TextBox And Left(Trim(ctl.Name), 4) = "txtC" Then ctl.BackColor = col
                Next
        
        Case "Product"
                For Each ctl In frm
                    If TypeOf ctl Is TextBox And Left(Trim(ctl.Name), 3) = "txt" Then ctl.BackColor = col
                    If TypeOf ctl Is ComboBox Then ctl.BackColor = col
                Next
        
    End Select
    
End Function

'**********Function to disable some buttons***********
Public Function fn_DISABLE_CONTROLS(frm As Form)

    With frm
        .cmdSave.Enabled = False
        .cmdCancel.Enabled = False
        
        .cmdEdit.Enabled = True
        .cmdDelete.Enabled = True
        .cmdAddNew.Enabled = True

        .cmdAdd.Enabled = False
        .cmdRemove.Enabled = False
    End With

End Function

'**********Function to enable some buttons***********
Public Function fn_ENABLE_CONTROLS(frm As Form)

    With frm
        .cmdSave.Enabled = True
        .cmdCancel.Enabled = True
        
        .cmdEdit.Enabled = False
        .cmdDelete.Enabled = False
        .cmdAddNew.Enabled = False

        .cmdAdd.Enabled = True
        .cmdRemove.Enabled = True
    End With

End Function

'Public Sub sub_GRID_ROW_MARKER(grd As MSFlexGrid)
'
'    With grd
'        If .Rows > 1 Then
'            If Trim(.TextMatrix(.Row, 1)) <> "" Then
'                If .TextMatrix(.Row, 0) = "" Then
'                    .TextMatrix(.Row, 0) = "ü"
'                    Else
'                        .TextMatrix(.Row, 0) = ""
'                End If
'                .col = 0
'                .CellFontName = "Wingdings"
'                .CellFontSize = "12"
'                .CellForeColor = vbBlue
'            End If
'        End If
'    End With
'
'End Sub

'Public Sub sub_ALIGN_TEXT_BOX_IN_GRID(grd As MSFlexGrid, txt As TextBox)
'
'    With grd
'        If .col > 1 And .Row > 0 Then
'            txt.Visible = True
'            txt.Move .CellLeft + .Left, .CellTop + .Top, .CellWidth, .CellHeight
'            txt.Text = .TextMatrix(.Row, .col)
'            .SetFocus
'        End If
'    End With
'
'    Call fn_HIGHLIGHT_TEXT(txt)
'
'End Sub

'**********Function to get the computer name*******************
Public Function fn_COMPUTER_NAME() As String

'    Dim Computer As String
'    Dim NameLen As Long
'
'    NameLen = 255
'    Computer = Space(NameLen)
'    GetComputerName Computer, NameLen
'
'    fn_COMPUTER_NAME = Left(Computer, NameLen)
    
End Function

'**********Function to select all in a listview*******************
Public Function fn_SELECT_ALL_IN_VIEW(lvw As ListView) As String
    Dim ctr As Long
    
    For ctr = 1 To lvw.ListItems.Count
        lvw.ListItems(ctr).Checked = True
    Next
    
End Function

'**********Function to unselect all in a listview*******************
Public Function fn_UNSELECT_ALL_IN_VIEW(lvw As ListView) As String
    Dim ctr As Long
    
    For ctr = 1 To lvw.ListItems.Count
        lvw.ListItems(ctr).Checked = False
    Next
    
End Function
