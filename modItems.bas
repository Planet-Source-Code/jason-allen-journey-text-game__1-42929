Attribute VB_Name = "modItems"
'declare all item types

Type ItemLantern
    Lit As Boolean
    Inv As Boolean
    Fuel As Integer
    FirstPickup As Boolean
End Type

Type ItemKey
    Inv As Boolean
    FirstPickup As Boolean
End Type

Type ItemDoor
    Locked As Boolean
    Open As Boolean
End Type

Public Type ItemMagazine
    Inv As Boolean
    FirstPickup As Boolean
End Type

Type ItemCamel
    Inv As Boolean
    FirstPickup As Boolean
End Type

Type ItemGarbage
    Inv As Boolean
    FirstPickup As Boolean
    Rats As Boolean
End Type

Type NPCRats
    Angry As Integer
    Guarding As Boolean
End Type

Function DefaultInventory()
    'set item defaults for game (items are declared in modGame)
    Lantern.Fuel = 20
    Lantern.Lit = False
    Lantern.Inv = False
    Lantern.FirstPickup = True
    Key.Inv = False
    Key.FirstPickup = True
    AlleyDoor.Locked = True
    AlleyDoor.Open = False
    Magazine.FirstPickup = True
    Magazine.Inv = False
    Camel.FirstPickup = True
    Camel.Inv = False
    Garbage.FirstPickup = True
    Garbage.Inv = False
    Garbage.Rats = False
    Rats.Angry = 0
    Rats.Guarding = True
    InventoryMax = 10
End Function

Function GetCarriedItems()
    'each item that can be in inventory (has a .inv type property) must be checked here
    Dim myString As String
    myString = ""
    'if item.inv then add the item to the carried list.  NOTE, that the name here must match the name
    'listed in the .dtf file and the corresponding item in the modGame module.
    If Lantern.Inv Then myString = myString & ", Lantern"
    If Key.Inv Then myString = myString & ", Key"
    If Magazine.Inv Then myString = myString & ", Magazine"
    If Camel.Inv Then myString = myString & ", Camel"
    If Garbage.Inv Then myString = myString & ", Garbage"
    
    'if there is anything in inventory, remove first two chars ", " that were auto added
    If Len(myString) > 2 Then myString = Right(myString, Len(myString) - 2)
    GetCarriedItems = myString
    
End Function

Function SaveItemStatus(ByVal FileName As String)
    'filename is same as .sav file but with .sai extension
    Open App.Path & "\" & FileName For Output As #8
    Print #8, InventoryCount 'send to file # of items in inventory
    'all Items from above types must be listed here with each of their corresponding properties
    Print #8, Lantern.FirstPickup & ":" & Lantern.Fuel & ":" & Lantern.Inv & ":" & Lantern.Lit
    Print #8, Key.FirstPickup & ":" & Key.Inv
    Print #8, AlleyDoor.Locked & ":" & AlleyDoor.Open
    Print #8, Magazine.FirstPickup & ":" & Magazine.Inv
    Print #8, Camel.FirstPickup & ":" & Camel.Inv
    Print #8, Garbage.FirstPickup & ":" & Garbage.Inv & ":" & Garbage.Rats
    Print #8, Rats.Angry & ":" & Rats.Guarding
    Close #8
End Function

Function LoadItemStatus(ByVal FileName As String)
    Dim FileRow As String
    Dim ItemArray() As String
    
    Open App.Path & "\" & FileName For Input As #9
    
    'get # of items in inventory
    Line Input #9, FileRow
    InventoryCount = FileRow
    
    'get all item statuses from file.  These must go in the same order as written in SaveItemStatus
    Line Input #9, FileRow
    ItemArray = Split(FileRow, ":")
    Lantern.FirstPickup = ItemArray(0): Lantern.Fuel = ItemArray(1): Lantern.Inv = ItemArray(2): Lantern.Lit = ItemArray(3)
    
    Line Input #9, FileRow
    ItemArray = Split(FileRow, ":")
    Key.FirstPickup = ItemArray(0): Key.Inv = ItemArray(1)
    
    Line Input #9, FileRow
    ItemArray = Split(FileRow, ":")
    AlleyDoor.Locked = ItemArray(0): AlleyDoor.Open = ItemArray(1)
    
    Line Input #9, FileRow
    ItemArray = Split(FileRow, ":")
    Magazine.FirstPickup = ItemArray(0): Magazine.Inv = ItemArray(1)
    
    Line Input #9, FileRow
    ItemArray = Split(FileRow, ":")
    Camel.FirstPickup = ItemArray(0): Camel.Inv = ItemArray(1)
    
    Line Input #9, FileRow
    ItemArray = Split(FileRow, ":")
    Garbage.FirstPickup = ItemArray(0): Garbage.Inv = ItemArray(1): Garbage.Rats = ItemArray(2)
    
    Line Input #9, FileRow
    ItemArray = Split(FileRow, ":")
    Rats.Angry = ItemArray(0): Rats.Guarding = ItemArray(1)
    
    Close #9
    
End Function

Function LanternDarkCheck()
    'check for lantern fuel (put out lantern if no fuel left), and then update location descriptions as necessary.
    If Lantern.Lit Then Lantern.Fuel = Lantern.Fuel - 1
    If Lantern.Fuel <= 0 And Lantern.Lit Then
        Lantern.Lit = False
        myBookmark = myRs.Bookmark
        myRs.MoveFirst: myRs.Find "Location='9.1'"
        myRs!Description = "It's Dark.  Except for a beam of light filtering in through the manhole above, you can't see anything."
        myRs.MoveFirst: myRs.Find "Location='9.2'"
        myRs!Description = "It's Dark.  You can't see anything."
        myRs.MoveFirst: myRs.Find "Location='9.3'"
        myRs!Description = "It's Dark.  You can't see anything, but you can hear the scratching of many sewer rats."
        myRs!Description2 = ""
        myRs.Bookmark = myBookmark
        frmJourney.txtActionPane.Text = "    You are out of Lantern oil.  Your lantern goes out." & vbCrLf & frmJourney.txtActionPane.Text
        UpdateCaptions
    End If
    'you die if you try to move in the dark
    If (myRs!Location = 9.2 Or myRs!Location = 9.3) And Lantern.Lit = False Then
        endgame = MsgBox("You trip over something and break your neck as you hit the ground.  The last thing you remember is a large sewer rat sitting on you and gnawing at your lower lip.", vbOKOnly, "Uh-oh")
        End
    End If
End Function


