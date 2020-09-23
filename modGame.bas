Attribute VB_Name = "modGame"
'declare items that have specific sub-attributes
Public Lantern As ItemLantern
Public Key As ItemKey
Public AlleyDoor As ItemDoor
Public Magazine As ItemMagazine
Public Camel As ItemCamel
Public Garbage As ItemGarbage
Public Rats As NPCRats

Function Game(ByVal Action As String)
    'Default Actions: N, E, S, W, U, D, Die, Sleep, Get, Drop
    'This map uses the extra actions: Light, Unlight, Open, Close, Lock, Unlock, Read, Look, Go
    Dim ActionList() As String
    If InStr(Action, " ") = 0 Then '1 word commands
        Select Case UCase(Action)
        'directions check for valid direction then call movement function and display result
        'default actions
            Case "N" 'North
                If myRs!N <> 0 Then
                    Call LanternDarkCheck
                    Call Travel("N"): Result = "You travel North"
                Else
                    Result = "You can't travel North"
                End If
            Case "E" 'East
                If myRs!e <> 0 Then
                    Call LanternDarkCheck
                    Call Travel("E"): Result = "You travel East"
                Else
                    Result = "You can't travel East"
                End If
            Case "S" 'South
                If myRs!S <> 0 Then
                    Call LanternDarkCheck
                    Call Travel("S"): Result = "You travel South"
                Else
                    Result = "You can't travel South"
                End If
            Case "W" 'West
                If myRs!W <> 0 Then
                    Call LanternDarkCheck
                    Call Travel("W"): Result = "You travel West"
                Else
                    Result = "You can't travel West"
                End If
            Case "D" 'Down
                If myRs!D <> 0 Then
                    Call LanternDarkCheck
                    Call Travel("D"): Result = "You climb Down"
                Else
                    Result = "You can't climb Down anything"
                End If
            Case "U" 'Up
                If myRs!U <> 0 Then
                    Call LanternDarkCheck
                    Call Travel("U"): Result = "You climb Up"
                Else
                    Result = "You can't climb Up anything"
                End If
            
            'useless functions, but fun
            Case "DIE"
                Result = "You try to commit suicide by slitting your wrists with your clothes... it doesn't work."
            Case "SLEEP"
                Result = "You take a nap... you wake up feeling refreshed"
            
            'actions with no item specified
            Case "DROP", "GET", "LIGHT", "UNLIGHT", "OPEN", "CLOSE", "LOCK", "UNLOCK", "READ"
                Result = Action & " what?"
            Case "LOOK", "GO"
                Result = Action & " where?"
            Case Else
                Result = "You can't do that"
        End Select
    
    Else '2+ word commands
        ActionList = Split(Action, " ", 2)
        Select Case UCase(ActionList(1))
            Case "LANTERN"
                Select Case UCase(ActionList(0))
                    Case "LIGHT"
                        If Lantern.Inv Then
                            If Lantern.Fuel <= 0 Then
                                Result = "You are out of Lantern oil"
                            Else
                                'update light-affected locations
                                Lantern.Lit = True
                                myBookmark = myRs.Bookmark
                                myRs.MoveFirst: myRs.Find "Location='9.1'"
                                myRs!Description = "You are in a sewer system.  The only light comes from the manhole above and your lantern."
                                myRs.MoveFirst: myRs.Find "Location='9.2'"
                                myRs!Description = "You are in a sewer system.  The only light you see comes from your lantern."
                                myRs.MoveFirst: myRs.Find "Location='9.3'"
                                myRs!Description = "You are at the end of the sewer system.  The only light you see comes from your lantern.  The sewer is full of rats."
                                If Key.FirstPickup Then myRs!Description2 = "  You see a small skeleton key on the ground."
                                myRs.Bookmark = myBookmark
                                UpdateCaptions
                                Result = "You light the Lantern"
                            End If
                        Else
                            Result = "You don't have a Lantern"
                        End If
                    
                    Case "UNLIGHT"
                        If Lantern.Inv Then
                            'update light-affected locations
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
                            UpdateCaptions
                            Result = "You unlight the Lantern"
                        Else
                            Result = "You don't have a Lantern"
                        End If
                    
                    Case "GET"
                        'call pickup lantern function and update status caption if necessary
                        If PickUp("Lantern", Lantern.Inv) Then
                            If Lantern.FirstPickup Then Lantern.FirstPickup = False: myRs!Description2 = "": UpdateCaptions
                        End If
                        
                    Case "DROP"
                        'call drop lantern function and update light-affected locations
                        If Drop("Lantern", Lantern.Inv) Then
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
                            UpdateCaptions
                        End If
                        
                    Case "LOOK"
                        'check the status of the lantern
                        If Lantern.Inv Then
                            If Lantern.Lit Then
                                Result = "The Lantern is lit - Fuel: " & Lantern.Fuel
                            Else
                                Result = "The Lantern is not lit - Fuel: " & Lantern.Fuel
                            End If
                        Else
                            Result = "You don't have a Lantern"
                        End If
                        
                    Case Else
                        Result = "You can't do that"
                End Select
                
            Case "KEY"
                Select Case UCase(ActionList(0))
                    Case "GET"
                        'call pickup key function and update status caption if necessary
                        'makes sure rats aren't guarding key.  if they are, they get angrier until you die
                        If myRs!Location = 9.3 And Key.FirstPickup Then
                            If Rats.Angry >= 2 Then
                                endgame = MsgBox("The rats attack!  The last thing you remember is a distant pain as one of the rats gnaws at a chunk of flesh hanging from your side.", vbOKOnly, "Game Over")
                                End
                            End If
                            If Rats.Guarding Then
                                Result = "The rats won't let you"
                                Rats.Angry = Rats.Angry + 1
                                GoTo NoKeyPickup
                            End If
                        End If
                        If PickUp("Key", Key.Inv) Then
                            If Key.FirstPickup Then Key.FirstPickup = False: myRs!Description2 = "": UpdateCaptions
                        End If
NoKeyPickup:

                    Case "DROP"
                        'call drop key function
                        If Drop("Key", Key.Inv) Then Key.Inv = False
                    
                    Case "LOOK"
                        'look at the key
                        If Key.Inv Then
                            Result = "It's a skeleton key"
                        Else
                            Result = "You don't have a key"
                        End If
                        
                    Case Else
                        Result = "You can't do that"
                End Select
                
            Case "DOOR"
                Select Case UCase(ActionList(0))
                    Case "UNLOCK"
                        'set door status to unlocked
                        If myRs!Location = 2.3 Then
                            If Key.Inv Then
                                If AlleyDoor.Locked And Not AlleyDoor.Open Then
                                    AlleyDoor.Locked = False
                                    myRs!Description3 = "  To the left you see an unlocked door."
                                    UpdateCaptions
                                    Result = "You unlock the door"
                                Else
                                    Result = "You can't do that"
                                End If
                            Else
                                Result = "You don't have a key"
                            End If
                        Else
                            Result = "What Door?"
                        End If
                        
                    Case "LOCK"
                        'set door status to locked
                        If myRs!Location = 2.3 Then
                            If Key.Inv Then
                                If Not AlleyDoor.Locked And Not AlleyDoor.Open Then
                                    AlleyDoor.Locked = True
                                    myRs!Description3 = "  To the left you see a locked door."
                                    UpdateCaptions
                                    Result = "You lock the door"
                                Else
                                    Result = "You can't do that"
                                End If
                            Else
                                Result = "You don't have a key"
                            End If
                        Else
                            Result = "What Door?"
                        End If
                    
                    Case "OPEN"
                        'set door status to open
                        If myRs!Location = 2.3 Then
                            If Not AlleyDoor.Locked Then
                                If Not AlleyDoor.Open Then
                                    AlleyDoor.Open = True
                                    myRs!Description3 = "  To the left you see an open door."
                                    UpdateCaptions
                                    Result = "You open the door"
                                Else
                                    Result = "It's already open"
                                End If
                            Else
                                Result = "It's locked"
                            End If
                        Else
                            Result = "What Door?"
                        End If
                    
                    Case "CLOSE"
                        'set door status to closed
                        If myRs!Location = 2.3 Then
                            If Not AlleyDoor.Open Then
                                Result = "It's already closed"
                            Else
                                AlleyDoor.Open = False
                                myRs!Description3 = "  To the left you see an unlocked door."
                                UpdateCaptions
                                Result = "You close the door"
                            End If
                        Else
                            Result = "What Door?"
                        End If
                        
                    Case "LOOK"
                        'look at the door to see if locked or not
                        If myRs!Location = 2.3 Then
                            If AlleyDoor.Locked Then
                                Result = "It's locked"
                            Else
                                Result = "It's unlocked"
                            End If
                        Else
                            Result = "What Door?"
                        End If
                        
                    Case "GO"
                        'enter the door (change locations)
                        If myRs!Location = 2.3 Then
                            If AlleyDoor.Open And Not AlleyDoor.Locked Then
                                myRs.Find "Location = '10.1'"
                                Result = "You enter the room.  The door shuts behind you and blends into the brick wall, leaving no trace of its existence"
                                UpdateCaptions
                            End If
                        Else
                            Result = "What Door?"
                        End If
                                        
                    Case Else
                        Result = "You can't do that"
                End Select
                
            Case "WINDOW"
                Select Case UCase(ActionList(0))
                    Case "GO", "LOOK", "OPEN", "CLOSE"
                        'you can't really do anything to the window
                        Result = "It's too high to reach"
                    
                    Case Else
                        Result = "You can't do that"
                End Select
                
            Case "MAGAZINE"
                Select Case UCase(ActionList(0))
                    Case "GET"
                        'call pickup magazine function and update status caption if necessary
                        If PickUp("Magazine", Magazine.Inv) Then
                            If Magazine.FirstPickup Then Magazine.FirstPickup = False: myRs!Description2 = "": UpdateCaptions
                        End If
                        
                    Case "DROP"
                        'call drop magazine function
                        If Drop("Magazine", Magazine.Inv) Then Magazine.Inv = False
                    
                    Case "LOOK"
                        'look at the magazine
                        If Magazine.Inv Then
                            Result = "It's an old dusty magazine.  You wonder what's inside?"
                        Else
                            Result = "You don't have a Magazine"
                        End If
                    
                    Case "READ"
                        'read the magazine
                        If Magazine.Inv Then
                            Result = "It's a copy of the latest Treasure Hunter's Digest.  Here's an interesting article: Priceless Golden Camel Lost - Large Reward.  Maybe you should try to find this Golden Camel."
                        Else
                            Result = "You don't have a Magazine"
                        End If
                    
                    Case Else
                        Result = "You can't do that"
                End Select
                
            Case "GARBAGE"
                Select Case UCase(ActionList(0))
                    Case "GET"
                        'pickup the garbage if the rats don't have it, update status cap if nec.
                        If Garbage.Rats Then
                            Result = "The Rats won't let you"
                        Else
                            If PickUp("Garbage", Garbage.Inv) Then
                                If Garbage.FirstPickup Then Garbage.FirstPickup = False: myRs!Description2 = "": UpdateCaptions
                            End If
                        End If
                        
                    Case "DROP"
                        'drop the garbage, check if rats are there, if so, they get it and don't guard anymore
                        If Drop("Garbage", Garbage.Inv) Then
                            If myRs!Location = 9.3 Then
                                Rats.Angry = 0
                                Rats.Guarding = False
                                Result = "You drop the Garbage.  The Rats grab the garbage and run to a corner and start fighting over the garbage."
                                Garbage.Rats = True
                            End If
                        End If
                        
                    Case "LOOK"
                        'look at the garbage
                        If Garbage.Inv Then
                            Result = "Yuck!  It Stinks!"
                        Else
                            Result = "You don't have the Garbage"
                        End If
                    
                    Case Else
                        Result = "You can't do that"
                End Select
            
            Case "GOLDEN CAMEL"
                Select Case UCase(ActionList(0))
                    Case "GET"
                        'call pickup camel function and update status caption if necessary
                        If PickUp("Golden Camel", Camel.Inv) Then
                            If Camel.FirstPickup Then Camel.FirstPickup = False: myRs!Description2 = "  Congratulations, Hero!  You've found the Golden Camel!  Get ready for other Journey adventures coming soon!": UpdateCaptions
                        End If
                        
                    Case "DROP"
                        'drop the camel
                        If Drop("Golden Camel", Camel.Inv) Then Camel.Inv = False
                    
                    Case "LOOK"
                        'look at the camel
                        If Camel.Inv Then
                            Result = "It's the Golden Camel!"
                        Else
                            Result = "You don't have the Golden Camel yet"
                        End If
                    
                    Case Else
                        Result = "You can't do that"
                End Select
            
            Case "RATS"
                If myRs!Location = 9.3 Then
                    Select Case UCase(ActionList(0))
                        Case "LOOK"
                            'look at the rats
                            If Garbage.Inv Then
                                Result = "The Rats eye your garbage greedily"
                            Else
                                If Garbage.Rats Then
                                    Result = "The Rats are fighting over the garbage"
                                Else
                                    Result = "The Rats look hungry"
                                End If
                            End If
                            
                        Case Else
                            'rats get angry if you try to do anything to them
                            If Rats.Angry >= 2 Then
                                endgame = MsgBox("The rats attack!  The last thing you remember is a distant pain as one of the rats gnaws at a chunk of flesh hanging from your side.", vbOKOnly, "Game Over")
                                End
                            End If
                            Result = "The Rats look angry"
                            Rats.Angry = Rats.Angry + 1
                    End Select
                Else
                    Result = "What Rats?"
                End If
                
            Case Else
                'item typed doesn't exist or can't be modified
                Result = "You can't do that"
        End Select
    End If
End Function

Function ShowHelpBox()
    Dim HelpString As String
    
    'show available commands and hints
    HelpString = "The following standard commands are available to you:" & vbCrLf & vbCrLf & _
        "N, E, S, W, U, D, Get, Drop, Die, Sleep" & vbCrLf & vbCrLf & vbCrLf & _
        "The following additional commands are available to you:" & vbCrLf & vbCrLf & "Go," & _
        " Lock, Unlock, Open, Close, Read, Look, Light, Unlight" & _
        vbCrLf & vbCrLf & vbCrLf & "Stuck?  Remember, not every item that you can interact " & _
        "with is listed in the 'items' summary.  Try being creative, like 'Go Window'"
    HelpBox = MsgBox(HelpString, vbOKOnly, "Common Commands and Hints")

End Function



