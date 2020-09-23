Attribute VB_Name = "modOpenSave"
Function OpenFile()
    Dim myPath As String
    Dim FileRow As String
    Dim myLocation As String
    Dim ItemList As String
    Dim NewCarrying() As String
    Dim Position As Integer
    Dim Field As String
    Dim ItemFile As String
    
    myPath = App.Path
    Dim CmDlg As New clsOpenSave
    CmDlg.DialogTitle = "Load Saved Journey"
    CmDlg.InitDir = myPath
    CmDlg.Filter = "Journey Save File (*.sav)|*.sav"
    On Error Resume Next 'Set up error handler
    CmDlg.CancelError = True
    CmDlg.ShowOpen
    If Err.Number = 32755 Then
        Result = MsgBox("Load Cancelled", vbOKOnly, "Load Cancelled")
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0 'Turn off error handler
    Open CmDlg.FileName For Input As #6
    
    'clear existing map recordset
    myRs.MoveFirst
    While myRs.EOF <> True
        myRs.Delete
        myRs.MoveNext
    Wend
    
    Do Until EOF(6)
        Line Input #6, FileRow
        'if "|" is in line, it is a map line
        If InStr(FileRow, Chr(124)) <> 0 Then
            With myRs
                .AddNew
                For Each fld In .Fields
                    ' If a | delimiter is found, field text is to the left of the delimiter.
                    If InStr(FileRow, Chr(124)) <> 0 Then
                        ' Move position to | delimiter.
                        Position = InStr(FileRow, Chr(124))
                        ' Assign field text to Field variable.
                        Field = Left(FileRow, Position - 1)
                    Else
                        ' If a | delimiter isn't found, field text is the
                        ' last field in the row.
                        Field = FileRow
                    End If
                    ' Strip off quotation marks.
                    If Left(Field, 1) = Chr(34) Then
                        Field = Left(Field, Len(Field) - 1)
                        Field = Right(Field, Len(Field) - 1)
                    End If

                    fld.Value = Field
                    If Field = "No Items" Then fld.Value = ""

                    ' Strip off field value text from text row.
                    FileRow = Right(FileRow, Len(FileRow) - Position)
                    Position = 0
                Next
                .Update
                .MoveFirst
            End With
        'if line begins with "*#" then it is the last location
        ElseIf InStr(FileRow, "*#") <> 0 Then
            FileRow = Right(FileRow, Len(FileRow) - 2)
            myLocation = RTrim(FileRow)
        End If
    Loop
    Close #6
    ItemFile = Left(RTrim(CmDlg.FileTitle), Len(RTrim(CmDlg.FileTitle)) - 2) & "i"
    'input all item statuses from .sai file
    Call LoadItemStatus(ItemFile)
    'go to last location
    myRs.Find "Location  = '" & myLocation & "'"
    UpdateCaptions
    frmJourney.txtAction.SetFocus
    frmJourney.txtActionPane.Text = "    Saved Game Loaded Successfully" & vbCrLf & frmJourney.txtActionPane.Text
    
End Function

Function SaveFile()
    Dim myPath As String
    Dim myBookmark
    Dim myLocation As String
    Dim ItemFile As String
    myRs.Update
    
    myPath = App.Path
    On Error GoTo 0
    Dim CmDlg As New clsOpenSave
    CmDlg.DialogTitle = "Save Journey"
    CmDlg.InitDir = myPath
    CmDlg.Filter = "Journey Save File (*.sav)|*.sav"
    On Error Resume Next 'Set up error handler
    CmDlg.CancelError = True
    CmDlg.ShowSave
    If Err.Number = 32755 Then
        Result = MsgBox("Save Cancelled", vbOKOnly, "Save Cancelled")
        Err.Clear
        Exit Function
    Else
        On Error GoTo 0
        'Process Save
        Open CmDlg.FileName For Output As #5
        myBookmark = myRs.Bookmark
        myLocation = myRs!Location
        myRs.MoveFirst
        While myRs.EOF <> True
            'print each line of current map recordset
            Print #5, RTrim(myRs!Location) & "|" & RTrim(myRs!Items) & "|" & RTrim(myRs!N) & "|" & RTrim(myRs!e) & "|" & RTrim(myRs!S) & "|" & RTrim(myRs!W) & "|" & RTrim(myRs!U) & "|" & RTrim(myRs!D) & "|" & RTrim(myRs!Description) & "|" & RTrim(myRs!Description2) & "|" & RTrim(myRs!Description3) & "|" & RTrim(myRs!Description4)
            myRs.MoveNext
        Wend
        myRs.Bookmark = myBookmark
        'print current location
        Print #5, vbCrLf & "*#" & myLocation
        Close #5
        ItemFile = Left(RTrim(CmDlg.FileTitle), Len(RTrim(CmDlg.FileTitle)) - 2) & "i"
        'save item statuses to .sai file
        Call SaveItemStatus(ItemFile)
    End If
    frmJourney.txtAction.SetFocus
    frmJourney.txtActionPane.Text = "    Game Saved Successfully" & vbCrLf & frmJourney.txtActionPane.Text
End Function
