Attribute VB_Name = "modUniversalActionFunctions"
Public InventoryCount As Integer
Public InventoryMax As Integer
Public Result As String 'holds result string that shows up in actionpane

Function Drop(ByVal Item As String, ByRef Inv As Boolean)
    Dim CurrentItems As String 'list of items in area that aren't in inventory
    Dim OneOnly As Boolean 'is there more than 1 item on ground after drop?

    CurrentItems = myRs!Items
    
    If Inv Then 'does item.inv = true (is item in inventory)
        If CurrentItems = "" Then OneOnly = True 'are there other items on ground?
        CurrentItems = CurrentItems & "; " & Item 'add item to list on ground
        If OneOnly Then CurrentItems = Right(CurrentItems, Len(CurrentItems) - 2) 'if only item on ground, remove the "; " that was auto added
        myRs!Items = CurrentItems 'update map recordset with current items on ground
        myRs.Update
        Result = Item & " dropped"
        InventoryCount = InventoryCount - 1 'decrease total items in inventory
        frmJourney.lblCarrying.Caption = GetCarriedItems() 'update captions
        frmJourney.lblItems.Caption = myRs!Items
        Drop = True 'function success
        Inv = False 'change item.inv to false
        UpdateCaptions
    Else
        Result = "You don't have " & Item
    End If
    
End Function

Function PickUp(ByVal Item As String, ByRef Inv As Boolean)
    Dim Found As Boolean 'is the item specified on the ground?
    Dim AvailableItems() As String 'items on the ground
    Dim NewItemsAvailable As String 'items on ground after pickup

    AvailableItems = Split(myRs!Items, "; ") 'list of what's on ground
    Found = False 'is item requested on ground?
    
    If Inv Then 'is item already in inventory?
        Result = "You already have " & Item
        Exit Function
    End If
    
    For i = 0 To UBound(AvailableItems) '0 to total items on ground
        If UCase(AvailableItems(i)) = UCase(Item) Then 'if item typed and item available matches...
            Found = True 'item found
            If InventoryCount < InventoryMax Then 'is inventory already full?
                InventoryCount = InventoryCount + 1 'add 1 to total inventory
                Result = AvailableItems(i) & " obtained"
                frmJourney.lblCarrying.Caption = GetCarriedItems()
                AvailableItems(i) = "" 'reset list on ground
                For j = 0 To UBound(AvailableItems) 'repopulate list on ground
                    If AvailableItems(j) <> "" Then NewItemsAvailable = NewItemsAvailable & AvailableItems(j) & "; "
                Next j
                If NewItemsAvailable <> "" Then NewItemsAvailable = Left(NewItemsAvailable, Len(NewItemsAvailable) - 2) 'pull last two characters ", "
                myRs!Items = NewItemsAvailable
                myRs.Update
                frmJourney.lblItems.Caption = myRs!Items
                PickUp = True ' function success
                Inv = True ' item.inv is now true
                UpdateCaptions
            Else
                Result = "You can't carry anymore" 'no more space available
            End If
        End If
    Next i
    
    If Not Found Then Result = Item & " is not in this area" 'not carrying and not found = not in area
    
End Function

Function Travel(ByVal myDirection As String)
    Dim NewLocation As String 'new coordinate from map recordset
    Select Case myDirection
        'move to location specified my map recordset
        Case "N"
            NewLocation = myRs!N
        Case "S"
            NewLocation = myRs!S
        Case "E"
            NewLocation = myRs!e
        Case "W"
            NewLocation = myRs!W
        Case "U"
            NewLocation = myRs!U
        Case "D"
            NewLocation = myRs!D
    End Select
    myRs.MoveFirst
    myRs.Find "Location = '" & NewLocation & "'"
    
    UpdateCaptions
    
End Function

Function UpdateCaptions()
    Dim AvailableDirs As String 'available directions
    AvailableDirs = ""
    'update whereareyou caption with all details from map recordset
    frmJourney.lblWhereAreYou.Caption = myRs!Description & myRs!Description2 & myRs!Description3 & myRs!Description4 'update captions
    'pull list of items currently in inventory
    frmJourney.lblCarrying.Caption = GetCarriedItems()
    'pull list of items on ground at location
    frmJourney.lblItems.Caption = myRs!Items
    'update available directions caption based on map recordset values
    If myRs!N <> 0 Then AvailableDirs = "North, "
    If myRs!e <> 0 Then AvailableDirs = AvailableDirs & "East, "
    If myRs!S <> 0 Then AvailableDirs = AvailableDirs & "South, "
    If myRs!W <> 0 Then AvailableDirs = AvailableDirs & "West, "
    If myRs!U <> 0 Then AvailableDirs = AvailableDirs & "Up, "
    If myRs!D <> 0 Then AvailableDirs = AvailableDirs & "Down, "
    'if any directions available, pull last two chars ", " from list
    If AvailableDirs <> "" Then AvailableDirs = Left(AvailableDirs, Len(AvailableDirs) - 2)
    frmJourney.lblDirection.Caption = AvailableDirs
End Function

