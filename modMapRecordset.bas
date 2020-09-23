Attribute VB_Name = "modMapRecordset"
Public myRs As ADODB.Recordset

Function MapSetup()
Dim fld As ADODB.Field
Dim Row As String
Dim Field As String
Dim Position As Integer

Set myRs = New ADODB.Recordset

'declare all fields for map recordset
With myRs
    .Fields.Append "Location", adChar, 5, adFldUpdatable
    .Fields.Append "Items", adChar, 1000, adFldUpdatable
    .Fields.Append "N", adChar, 5, adFldUpdatable
    .Fields.Append "E", adChar, 5, adFldUpdatable
    .Fields.Append "S", adChar, 5, adFldUpdatable
    .Fields.Append "W", adChar, 5, adFldUpdatable
    .Fields.Append "U", adChar, 5, adFldUpdatable
    .Fields.Append "D", adChar, 5, adFldUpdatable
    .Fields.Append "Description", adChar, 1000, adFldUpdatable
    .Fields.Append "Description2", adChar, 1000, adFldUpdatable
    .Fields.Append "Description3", adChar, 1000, adFldUpdatable
    .Fields.Append "Description4", adChar, 1000, adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With

myPath = App.Path
Open myPath & "\Journey.dtf" For Input As #1

Line Input #1, Row
Do Until EOF(1)
    Line Input #1, Row
    With myRs
        .AddNew
        For Each fld In .Fields
            ' If a | delimiter is found, field text is to the left of the delimiter.
            If InStr(Row, Chr(124)) <> 0 Then
               ' Move position to | delimiter.
               Position = InStr(Row, Chr(124))
               ' Assign field text to Field variable.
               Field = Left(Row, Position - 1)
            Else
               ' If a | delimiter isn't found, field text is the
               ' last field in the row.
               Field = Row
            End If

            ' Strip off quotation marks.
            If Left(Field, 1) = Chr(34) Then
               Field = Left(Field, Len(Field) - 1)
               Field = Right(Field, Len(Field) - 1)
            End If

            fld.Value = Field
            If Field = "No Items" Then fld.Value = ""

            ' Strip off field value text from text row.
            Row = Right(Row, Len(Row) - Position)
            Position = 0
        Next
        .Update
        .MoveFirst
    End With
Loop
Close #1
End Function
