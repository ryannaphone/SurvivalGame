Attribute VB_Name = "MiscSubs"

Sub LogGather()
    FailCount = 0
    Call UpdateEnergy(-8, False)
    If Axe = True Then
        LogsPerGather = 6
    Else
        LogsPerGather = 4
    End If
    TRN = RandomInteger(1, 4)
    For i = 1 To 4
        RangeS = 4
        RangeE = 5
        RangeS = RangeS - (1 * i)
        RangeE = RangeE - (1 * i)
        If TRN > RangeS And TRN <= RangeE Then
            Call UpdateLog(LogsPerGather - i, False)
            MainForm.NewsLabel.Caption = "News: You found " & (LogsPerGather - i) & " logs!"
        Else
            FailCount = FailCount + 1
        End If
    Next i
    If FailCount = 3 Then MainForm.NewsLabel.Caption = "News: You failed to find a log."
End Sub

Sub CraftAxe()
    If Axe = True Then
        MainForm.NewsLabel.Caption = "News: You already have an axe."
        Exit Sub
    End If
    If Sticks >= 1 And Stones >= 2 Then
        Axe = True
        Call UpdateStick(-1, False)
        Call UpdateStone(-2, False)
        Call UpdateInventory(3, False)
        MainForm.AxeLabel.Caption = "Axe"
        MainForm.NewsLabel.Caption = "News: You crafted an axe."
        Call UpdateThingsCrafted(1, False)
    Else
        MainForm.NewsLabel.Caption = "News: You don't have the required components."
    End If
End Sub

Sub CraftKnife()
    If Knife = True Then
        MainForm.NewsLabel.Caption = "News: You already have a knife."
        Exit Sub
    End If
    If Sticks >= 1 And Stones >= 1 Then
        Knife = True
        Call UpdateStick(-1, False)
        Call UpdateStone(-1, False)
        Call UpdateInventory(-1, False)
        MainForm.KnifeLabel.Caption = "Knife"
        MainForm.NewsLabel.Caption = "News: You crafted a knife."
        Call UpdateThingsCrafted(1, False)
    Else
        MainForm.NewsLabel.Caption = "News: You don't have the required components."
    End If
End Sub

Sub BerryGather()
    FailCount = 0
    TRN = RandomInteger(GatheringLvl, 20)
    Temp1 = 20
    Call UpdateEnergy(-5, False)
    For i = 1 To 10
        RangeS = 20
        RangeE = 22
        RangeS = RangeS - (2 * i)
        RangeE = RangeE - (2 * i)
        If TRN > RangeS And TRN <= RangeE Then
            Call UpdateBerry(Temp1 - (i * 2), False)
            MainForm.NewsLabel.Caption = "News: You found " & (Temp1 - (i * 2)) & " berries!"
        Else
            FailCount = FailCount + 1
        End If
    Next i
    If FailCount = 10 Then MainForm.NewsLabel.Caption = "News: You failed to find a berry."
End Sub

Sub LeafGather()
    FailCount = 0
    TRN = RandomInteger(GatheringLvl, 20)
    Temp1 = 30
    Call UpdateEnergy(-2, False)
    For i = 1 To 10
        RangeS = 20
        RangeE = 22
        RangeS = RangeS - (2 * i)
        RangeE = RangeE - (2 * i)
        If TRN > RangeS And TRN <= RangeE Then
            Call UpdateLeaf(Temp1 - (2 * i), False)
            MainForm.NewsLabel.Caption = "News: You found " & (Temp1 - (2 * i)) & " leaves!"
        Else
            FailCount = FailCount + 1
        End If
    Next i
    If FailCount = 10 Then MainForm.NewsLabel.Caption = "News: You failed to find a leaf."
End Sub

Sub StickGather()
    FailCount = 0
    TRN = RandomInteger(GatheringLvl, 20)
    Temp1 = 10
    Call UpdateEnergy(-5, False)
    For i = 1 To 10
        RangeS = 20
        RangeE = 22
        RangeS = RangeS - (2 * i)
        RangeE = RangeE - (2 * i)
        If TRN > RangeS And TRN <= RangeE Then
            Call UpdateStick(Temp1 - (i), False)
            MainForm.NewsLabel.Caption = "News: You found " & (Temp1 - (i)) & " sticks!"
        Else
            FailCount = FailCount + 1
        End If
    Next i
    If FailCount = 10 Then MainForm.NewsLabel.Caption = "News: You failed to find a stick."
End Sub

Sub WaterGather()
    FailCount = 0
    TRN = RandomInteger(GatheringLvl, 20)
    Temp1 = 10
    Call UpdateEnergy(-5, False)
    For i = 1 To 10
        RangeS = 20
        RangeE = 22
        RangeS = RangeS - (2 * i)
        RangeE = RangeE - (2 * i)
        If TRN > RangeS And TRN <= RangeE Then
            Call UpdateWater(Temp1 - (i), False)
            MainForm.NewsLabel.Caption = "News: You found " & (Temp1 - (i)) & " cups of water!"
        Else
            FailCount = FailCount + 1
        End If
    Next i
    If FailCount = 10 Then MainForm.NewsLabel.Caption = "News: You failed to find any water."
End Sub

Sub StoneGather()
    FailCount = 0
    TRN = RandomInteger(GatheringLvl, 20)
    Temp1 = 10
    Call UpdateEnergy(-5, False)
    For i = 1 To 10
        RangeS = 20
        RangeE = 22
        RangeS = RangeS - (2 * i)
        RangeE = RangeE - (2 * i)
        If TRN > RangeS And TRN <= RangeE Then
            Call UpdateStone(Temp1 - (i), False)
            MainForm.NewsLabel.Caption = "News: You found " & (Temp1 - (i)) & " stones!"
        Else
            FailCount = FailCount + 1
        End If
    Next i
    If FailCount = 10 Then MainForm.NewsLabel.Caption = "News: You failed to find a stones."
End Sub

Sub ConsumeBerry()
    'TRN = RandomInteger(1, 10)
    If RandomInteger(1, 10) > 4 Then Call UpdateHunger(1, False)
    If RandomInteger(1, 10) > 8 Then Call UpdateEnergy(1, False)
    If RandomInteger(1, 10) > 9 Then Call UpdateThirst(1, False)
    If RandomInteger(1, 10) > 9 Then Call UpdateHealth(1, False)
    Call UpdateBerry(-1, False)
End Sub

Sub ConsumeMeat()
    TRN = RandomInteger(1, 3)
    If TRN = 1 Then Call UpdateHunger(5, False)
    If TRN = 2 Then Call UpdateHunger(6, False)
    If TRN = 3 Then Call UpdateHunger(7, False)
    TRN = RandomInteger(1, 3)
    If TRN = 1 Then Call UpdateEnergy(3, False)
    If TRN = 2 Then Call UpdateEnergy(4, False)
    If TRN = 3 Then Call UpdateEnergy(5, False)
    If RandomInteger(1, 10) > 9 Then Call UpdateThirst(1, False)
    TRN = RandomInteger(1, 10)
    If TRN > 7 Then Call UpdateHealth(1, False)
    If TRN = 10 Then Call UpdateHealth(2, False)
    Call UpdateMeat(-1, False)
End Sub

Sub ConsumeRawMeat()
    TRN = RandomInteger(1, 3)
    If TRN = 1 Then Call UpdateHunger(1, False)
    If TRN = 2 Then Call UpdateHunger(2, False)
    If TRN = 3 Then Call UpdateHunger(3, False)
    TRN = RandomInteger(1, 2)
    If TRN = 1 Then Call UpdateEnergy(1, False)
    If TRN = 2 Then Call UpdateEnergy(2, False)
    If RandomInteger(1, 10) > 9 Then Call UpdateThirst(1, False)
    TRN = RandomInteger(1, 10)
    If TRN > 4 Then Call UpdateHealth(-2, False)
    Call UpdateRawMeat(-1, False)
End Sub

Sub CookRawMeat()
    Temp1 = 9
    TRN = RandomInteger(CookingLvl, 20)
    Call UpdateRawMeat(-1, False)
    If TRN > 9 Then
        Call UpdateMeat(1, False)
        MainForm.NewsLabel.Caption = "News: You cooked some meat."
    Else
        MainForm.NewsLabel.Caption = "News: You burnt some meat."
    End If
End Sub

Sub HuntRabbit()
    If RandomInteger(1, 3) = 3 And WolfPrompt1 = False Then
        Load SituationForm
        SituationForm.Show
        SituationForm.SituationInfo.Caption = "Situation Info: While hunting, a sickly looking wolf attacks the rabbit you are tracking. What do you do?"
        SituationForm.Option1.Caption = "Try to kill them both."
        SituationForm.Option2.Caption = "Leave them alone."
        WolfPrompt1 = True
        WolfPrompt2 = False
        Exit Sub
    End If
    TRN = RandomInteger(HuntingLvl, 20)
    Call UpdateEnergy(-6, False)
    If Spear = False Then
        Temp1 = 17
    Else
        Temp1 = 15
    End If
    If Bow = True And Arrows >= 5 Then
        Temp1 = Temp1 - 2
        Call UpdateArrows(-RandomInteger(0, 2), False)
    End If
    If TRN >= Temp1 Then
        Call UpdateRawMeat(1, False)
        MainForm.NewsLabel.Caption = "News: You killed a rabbit!"
        If Amount > 0 Then Call UpdateAnimalsHunted(2, False)
        If Knife = True Then
            Call UpdateFurs(RandomInteger(1, 3), False)
        End If
    Else
        If RandomInteger(1, 10) > 3 Then
            If Amount > 0 Then Call UpdateAnimalsHunted(1, False)
        End If
        MainForm.NewsLabel.Caption = "News: You failed in your hunt."
    End If
End Sub

Sub CraftSpear()
    If Spear = True Then
        MainForm.NewsLabel.Caption = "News: You already have a spear."
        Exit Sub
    End If
    If Sticks >= 2 And Stones >= 2 Then
        Spear = True
        MainForm.SpearLabel.Caption = "Spear"
        MainForm.NewsLabel.Caption = "News: You crafted a spear."
        Call UpdateThingsCrafted(1, False)
    Else
        MainForm.NewsLabel.Caption = "News: You don't have the required components."
    End If
End Sub

Sub CraftFire()
    If Fire = True Then
        MainForm.NewsLabel.Caption = "News: You already have a fire."
        Exit Sub
    End If
    If Sticks >= 10 And Logs >= 3 And Leaves >= 25 Then
        Fire = True
        Call UpdateStick(-10, False)
        Call UpdateLog(-3, False)
        Call UpdateLeaf(-25, False)
        FireStartDay = DayNum
        MainForm.FireLabel.Caption = "Fire"
        MainForm.NewsLabel.Caption = "News: You made a fire."
        Call UpdateThingsCrafted(1, False)
    End If
End Sub

Sub BuildSShelter()
    If StickShelt = True Then
        MainForm.NewsLabel.Caption = "News: You already have a stick shelter."
        Exit Sub
    End If
    If Sticks >= 15 And Leaves >= 50 Then
        StickShelt = True
        UpdateEnergyPerNight
        MainForm.NewsLabel.Caption = "News: You built a stick shelter."
        Call UpdateStick(-15, False)
        Call UpdateLeaf(-50, False)
        MainForm.SSLabel.Caption = "Stick Shelter"
    End If
End Sub

Sub BuildCabin()
    If Cabin = True Then
        MainForm.NewsLabel.Caption = "News: You already have a cabin."
        Exit Sub
    End If
    If Logs >= 50 And Sticks >= 50 Then
        Cabin = True
        UpdateEnergyPerNight
        MainForm.NewsLabel.Caption = "News: You built a log cabin."
        Call UpdateStick(-50, False)
        Call UpdateLog(-50, False)
        MainForm.CabinLabel.Caption = "Cabin"
    End If
End Sub

Sub BuildTrap()
    If Stones >= 2 And Sticks >= 5 And Berries >= 2 Then
        Call UpdateTraps(1, False)
        MainForm.NewsLabel.Caption = "News: You built a rabbit trap."
        Call UpdateStick(-5, False)
        Call UpdateStone(-2, False)
        Call UpdateBerry(-2, False)
        'ReDim TrapsSet(1 To Traps)
        'TrapsSet(Traps) = True
        SetTraps = SetTraps + 1
    Else
        MainForm.NewsLabel.Caption = "News: You lack the required components."
    End If
End Sub

Sub BuildBed()
    If Bed = True Then
        MainForm.NewsLabel.Caption = "News: You already have a bed."
        Exit Sub
    End If
    If Logs >= 6 And Sticks >= 20 And Furs >= 25 Then
        If Cabin = True Then
            Bed = True
            UpdateEnergyPerNight
            MainForm.NewsLabel.Caption = "News: You built a bed."
            Call UpdateStick(-20, False)
            Call UpdateLog(-6, False)
            Call UpdateFurs(-25, False)
            MainForm.BedLabel.Caption = "Bed"
        Else
            MainForm.NewsLabel.Caption = "News: You need to build a cabin before building a bed."
        End If
    Else
        MainForm.NewsLabel.Caption = "News: You lack the required components."
    End If
End Sub

Sub CraftSticks()
    If Logs >= 1 Then
        If Knife = False Then
            Temp1 = 4
        Else
            Temp1 = 8
        End If
        Call UpdateLog(-1, False)
        Call UpdateStick(Temp1, False)
        MainForm.NewsLabel.Caption = "News: You crafted some sticks."
        Call UpdateThingsCrafted(1, False)
    Else
        MainForm.NewsLabel.Caption = "News: You need a log."
    End If
End Sub

Sub ConsumeWater()
    If Water > 0 Then
        Call UpdateWater(-1, False)
        TRN = RandomInteger(1, 10)
        If TRN > 7 Then Call UpdateThirst(2, False)
        If TRN <= 7 And TRN > 2 Then Call UpdateThirst(1, False)
        TRN = RandomInteger(1, 10)
        If TRN > 7 Then Call UpdateEnergy(1, False)
    Else
        MainForm.NewsLabel.Caption = "News: You need water in order to drink water..."
    End If
End Sub

Sub RestartGame()
    Spear = False
    MainForm.SpearLabel.Caption = "No Spear"
    Fire = False
    MainForm.FireLabel.Caption = "No Fire"
    Axe = False
    MainForm.AxeLabel.Caption = "No Axe"
    StickShelt = False
    MainForm.SSLabel.Caption = "No Stick Shelter"
    Knife = False
    MainForm.KnifeLabel.Caption = "No Knife"
    Bed = False
    MainForm.BedLabel.Caption = "No Bed"
    Cabin = False
    MainForm.CabinLabel.Caption = "No Cabin"
    ConsoleOpen = True
    SnipeEnabled = False
    WolfPrompt1 = False
    WolfPrompt2 = True
    WolfPrompt3 = True
    UpdateConsoleOpen
    UpdateDay (True)
    Call UpdateEnergyPerNight
    Call UpdateRabbitCage(True)
    Call UpdateBow(True)
    Call UpdateHealth(0, True)
    Call UpdateEnergy(0, True)
    Call UpdateHunger(0, True)
    Call UpdateThirst(0, True)
    Call UpdateBerry(0, True)
    Call UpdateStick(0, True)
    Call UpdateWater(0, True)
    Call UpdateLog(0, True)
    Call UpdateLeaf(0, True)
    Call UpdateStone(0, True)
    Call UpdateFurs(0, True)
    Call UpdateInventory(0, True)
    Call UpdateMeat(0, True)
    Call UpdateRawMeat(0, True)
    Call UpdateRabbits(0, True)
    Call UpdateArrows(0, True)
    Call UpdateVines(0, True)
    Call UpdateFeathers(0, True)
    Call UpdateTraps(0, True)
    Call UpdateBirdTraps(0, True)
    Call UpdateLiveTraps(0, True)
    Call UpdateAllLevels(0, True)
    Call UpdateThingsCooked(0, True)
    Call UpdateThingsCrafted(0, True)
    Call UpdateEnergySpent(0, True)
    Call UpdateAllCollected(0, True)
    Call UpdateAnimalsHunted(0, True)
    Call UpdateHealthGain(0, True)
    Call UpdateActionsCompleted(0, True)
    Call UpdateCWolf(True)
    Call UpdateCWolfXP(0, True)
    Call UpdateCWolfHealth(0, True)
    Call UpdateCWolfHunger(0, True)
    Call UpdateCWolfBond(0, True)
    Call UpdateWolfBond(0, False)
    Call UpdateRabbitHunger(0, True)
    MainForm.InfoLabel.Caption = "Info: Enjoy the game!"
    TabSelect = True
    MainForm.BackButton.Visible = False
    MainForm.GRFrame.Visible = False
    MainForm.TSFrame.Visible = True
    SelectedSlot = 0
End Sub

Sub EndDay()
    Call UpdateEnergyPerNight
    Call UpdateHunger(-5, False)
    Call UpdateThirst(-5, False)
    If Rabbits > 0 Then Call UpdateRabbitHunger(-1 * Rabbits, False)
    If CWolf = True Then Call UpdateCWolfHunger(-5, False)
    CWolfDailyAction = False
    CheckFireTime
    TRN = RandomInteger(1, 100)
    If TRN < (4 / 3) * (Energy + Hunger + Thirst) Then Call UpdateHealth(1, False)
    Call UpdateEnergy(EnergyPerNight, False)
    UpdateDay (False)
    If Traps >= 1 Then CheckTraps
    If BTraps >= 1 Then CheckBirdTraps
    If LTraps >= 1 Then CheckLiveTraps
    MainForm.InfoLabel.Caption = "Info: Enjoy the game!"
    MainForm.NewsLabel.Caption = "News: Have a wonderful Day# " & DayNum & "!"
    If RandomInteger(1, 3) = 3 And WolfPrompt3 = False Then
        Load SituationForm
        SituationForm.Show
        SituationForm.SituationInfo.Caption = "Situation Info: You wake up to find the sickly looking wolf sitting near you. What do you do?"
        SituationForm.Option1.Caption = "Attack it."
        SituationForm.Option2.Caption = "Offer it some meat. (Requires 1 raw or cooked meat)"
        WolfPrompt3 = True
    End If
    If DayNum >= 10 Then
        If RawMeat + Meat < 20 Then
            TRN = RandomInteger(RawMeat + Meat, 20)
            If TRN >= 15 Then Call BearAttack
        Else
            If RandomInteger(1, 10) <= 9 Then Call BearAttack
        End If
        If Rabbits > 0 And Rabbits < 20 Then
            TRN = RandomInteger(Rabbits, 20)
            If TRN >= 15 Then Call WolfAttack
        ElseIf Rabbits > 20 Then
            If RandomInteger(1, 10) <= 9 Then Call WolfAttack
        End If
    End If
    If Rabbits >= 2 Then
        RabbitMating
    End If
    If AutoSave = 2 Then
        Call SaveGame(SelectedSlot)
    End If
End Sub

Sub CheckTraps()
    Temp1 = SetTraps
    For i = 1 To Temp1
        TRN = RandomInteger(1, 10)
        If TRN < 3 Then
            Call UpdateTraps(-1, False)
            MsgBox "One of your rabbit traps broke."
            SetTraps = SetTraps - 1
        End If
        If TRN >= 7 Then
            If RandomInteger(1, 3) = 3 And WolfPrompt2 = False Then
                Load SituationForm
                SituationForm.Show
                SituationForm.SituationInfo.Caption = "Situation Info: While checking your traps, you find a sickly looking wolf eating a rabbit that was caught in one of your traps. As you walk up, it turns around and looks at you. What do you do?"
                SituationForm.Option1.Caption = "Offer it some meat. (Requires 1 raw or cooked meat)"
                SituationForm.Option2.Caption = "Throw a rock at it."
                WolfPrompt2 = True
                WolfPrompt3 = False
                SetTraps = SetTraps - 1
            Else
                Call UpdateRawMeat(1, False)
                If Knife = True Then Call UpdateFurs(RandomInteger(1, 3), False)
                MsgBox "One of your traps caught a rabbit! It will need to be reset to catch another."
                SetTraps = SetTraps - 1
            End If
        End If
    Next i
End Sub

Sub MaintainTraps()
    If SetTraps = Traps Then
        MainForm.NewsLabel.Caption = "News: All of your traps have berries."
        Exit Sub
    End If
    If Berries >= 2 Then
        Call UpdateBerry(-2, False)
        MainForm.NewsLabel.Caption = "News: You put more berries in a trap."
        SetTraps = SetTraps + 1
        Exit Sub
    Else
        MainForm.NewsLabel.Caption = "News: You don't have enough berries."
        Exit Sub
    End If
End Sub

Sub MaintainFire()
    If Logs >= 2 Then
        If FireStartDay <> DayNum Then
            Call UpdateLog(-2, False)
            'Call UpdateStick(-5, False)
            FireStartDay = DayNum
            MainForm.NewsLabel.Caption = "News: You added fuel to your fire."
        Else
            MainForm.NewsLabel.Caption = "News: You just added fuel."
        End If
    Else
        MainForm.NewsLabel.Caption = "News: You lack the required components."
    End If
End Sub

Sub CheckFireTime()
    If Fire = True Then
        If DayNum - FireStartDay = 5 Then
            Fire = False
            MsgBox "Your fire went out."
        End If
    End If
End Sub

Sub HuntDeer()
    TRN = RandomInteger(HuntingLvl, 20)
    Call UpdateEnergy(-8, False)
    If Spear = False Then
        Temp1 = 19
    Else
        Temp1 = 17
    End If
    If Bow = True And Arrows >= 5 Then
        Temp1 = Temp1 - 2
        Call UpdateArrows(-RandomInteger(0, 3), False)
    End If
    If Knife = True Then
        Call UpdateFurs(RandomInteger(2, 4), False)
    End If
    'Call UpdateEnergy(-5, False)
    If TRN >= Temp1 Then
        Call UpdateRawMeat(4, False)
        MainForm.NewsLabel.Caption = "News: You killed a deer!"
        If Amount > 0 Then Call UpdateAnimalsHunted(2, False)
    Else
        If RandomInteger(1, 10) > 3 Then
            If Amount > 0 Then Call UpdateAnimalsHunted(1, False)
        End If
        MainForm.NewsLabel.Caption = "News: You failed in your hunt."
    End If
End Sub

Sub CWolfHunt()
    TRN = RandomInteger(CWolfLvl, 20)
    Temp1 = 15
    CWolfDailyAction = True
    'Call UpdateEnergy(-5, False)
    If TRN >= Temp1 Then
        Call UpdateRawMeat(1, False)
        MainForm.NewsLabel.Caption = "News: Your wolf brought you some meat."
        Call UpdateCWolfXP(3, False)
        Call UpdateCWolfHunger(5, False)
    Else
        If RandomInteger(1, 10) > 3 Then
            Call UpdateCWolfXP(1, False)
        End If
        MainForm.NewsLabel.Caption = "News: You wolf returned empty handed."
    End If
End Sub

Sub WolfFeed()
    If RawMeat >= 1 Then
        Call UpdateRawMeat(-1, False)
        Call UpdateCWolfHunger(5, False)
        Call UpdateCWolfBond(5, False)
        MainForm.NewsLabel.Caption = "News: You fed your wolf some raw meat."
    ElseIf Meat >= 1 Then
        Call UpdateMeat(-1, False)
        Call UpdateCWolfHunger(5, False)
        Call UpdateCWolfBond(5, False)
        MainForm.NewsLabel.Caption = "News: You fed your wolf some meat."
    End If
End Sub

Sub WolfPet()
    If Energy >= 3 Then
        MainForm.NewsLabel.Caption = "News: You showed your wolf some affection."
        Call UpdateEnergy(-3, False)
        Call UpdateCWolfBond(2, False)
    Else
        MainForm.NewsLabel.Caption = "News: You don't have enough energy to complete this action."
    End If
End Sub

Sub BearAttack()
    Temp1 = 10
    If Spear = True Then Temp1 = Temp1 - 2
    If Bow = True And Arrows >= 5 Then Temp1 = Temp1 - 2
    If StrengthLvl >= 5 Then Temp1 = Temp1 - 1
    If HuntingLvl >= 5 Then Temp1 = Temp1 - 1
    TRN = RandomInteger(1, Temp1)
    If TRN = 10 Then
        Call UpdateHealth(-RandomInteger(10, 50), False)
        MsgBox "A bear attacked you and seriously injured you!"
    End If
    If TRN > 7 And TRN <> 10 Then
        Call UpdateHealth(-RandomInteger(5, 20), False)
        MsgBox "A bear attacked you and injured you!"
    End If
    If TRN > 3 And TRN < 7 Then
        If RawMeat > 0 Then
            Call UpdateRawMeat(-RandomInteger(1, RawMeat), False)
            MsgBox "A bear attacked you and stole some raw meat!"
        ElseIf Meat > 0 Then
            Call UpdateMeat(-RandomInteger(1, Meat), False)
            MsgBox "A bear attacked you and stole some cooked meat!"
        End If
    End If
    If TRN < 3 Then
        MsgBox "A bear attacked you, but not much happened."
    End If
End Sub

Sub CraftArrows()
    If Sticks >= 5 And Stones >= 5 And Feathers >= 5 Then
        Call UpdateArrows(5, False)
        Call UpdateStick(-5, False)
        Call UpdateStone(-5, False)
        Call UpdateFeathers(-5, False)
        MainForm.NewsLabel.Caption = "News: You crafted 5 arrows."
    Else
        MainForm.NewsLabel.Caption = "News: You lack the required components."
    End If
End Sub

Sub CraftBow()
    If Bow = False Then
        If Sticks >= 3 And Vines >= 2 Then
            Call UpdateBow(False)
            Call UpdateStick(-3, False)
            Call UpdateVines(-2, False)
            Call UpdateInventory(-2, False)
            Bow = True
            MainForm.NewsLabel.Caption = "News: You crafted a bow!"
        Else
            MainForm.NewsLabel.Caption = "News: You lack the required components."
        End If
    Else
        MainForm.NewsLabel.Caption = "News: You already have a bow."
    End If
End Sub

Sub HuntBear()
    TRN = RandomInteger(HuntingLvl, 20)
    Call UpdateEnergy(-10, False)
    If Spear = False Then
        Temp1 = 21
    Else
        Temp1 = 19
    End If
    If Bow = True And Arrows >= 5 Then
        Temp1 = Temp1 - 2
        Call UpdateArrows(-RandomInteger(0, 5), False)
    End If
    If Knife = True Then
        Call UpdateFurs(RandomInteger(3, 10), False)
    End If
    If TRN >= Temp1 Then
        Call UpdateRawMeat(10, False)
        MainForm.NewsLabel.Caption = "News: You killed a bear!"
        If Amount > 0 Then Call UpdateAnimalsHunted(5, False)
    ElseIf TRN <= Temp1 - 10 Then
        If RandomInteger(1, 10) > 3 Then
            If Amount > 0 Then Call UpdateAnimalsHunted(1, False)
        End If
        MainForm.NewsLabel.Caption = "News: You failed in your hunt."
    ElseIf TRN > Temp1 - 10 And TRN < Temp1 Then
        Call UpdateHealth(-25, False)
        MainForm.NewsLabel.Caption = "News: The bear attacked you."
    End If
End Sub

Sub HuntWolves()
    TRN = RandomInteger(HuntingLvl, 20)
    Call UpdateEnergy(-9, False)
    If Spear = False Then
        Temp1 = 20
    Else
        Temp1 = 17
    End If
    If Bow = True And Arrows >= 5 Then
        Temp1 = Temp1 - 2
        Call UpdateArrows(-RandomInteger(0, 4), False)
    End If
    If Knife = True Then
        Call UpdateFurs(RandomInteger(3, 5), False)
    End If
    'Call UpdateEnergy(-5, False)
    If TRN >= Temp1 Then
        Call UpdateRawMeat(5, False)
        MainForm.NewsLabel.Caption = "News: You killed a wolf!"
        If Amount > 0 Then Call UpdateAnimalsHunted(3, False)
    ElseIf TRN <= Temp1 - 10 Then
        If RandomInteger(1, 10) > 3 Then
            If Amount > 0 Then Call UpdateAnimalsHunted(1, False)
        End If
        MainForm.NewsLabel.Caption = "News: You failed in your hunt."
    ElseIf TRN > Temp1 - 10 And TRN < Temp1 Then
        Call UpdateHealth(-25, False)
        MainForm.NewsLabel.Caption = "News: The wolf attacked you."
    End If
End Sub

Sub WolfAttack()
    Temp1 = 10
    If Spear = True Then Temp1 = Temp1 - 2
    If Bow = True And Arrows >= 5 Then Temp1 = Temp1 - 2
    If StrengthLvl >= 5 Then Temp1 = Temp1 - 1
    If HuntingLvl >= 5 Then Temp1 = Temp1 - 1
    TRN = RandomInteger(1, Temp1)
    If TRN = 10 Then
        Call UpdateHealth(-RandomInteger(10, 25), False)
        MsgBox "A wolf attacked you and seriously injured you!"
    End If
    If TRN > 7 And TRN <> 10 Then
        Call UpdateHealth(-RandomInteger(5, 15), False)
        MsgBox "A wolf attacked you and injured you!"
    End If
    If TRN > 3 And TRN < 7 Then
        If RawMeat > 0 Then
            Call UpdateRabbits(-RandomInteger(1, Rabbits), False)
            MsgBox "A wolf attacked you and killed some rabbits!"
        End If
    End If
    If TRN < 3 Then
        MsgBox "A wolf attacked you, but not much happened."
    End If
End Sub

Sub BuildRabbitCage()
    If Sticks >= 25 And Stones >= 10 Then
        Call UpdateStick(-25, False)
        Call UpdateStone(-10, False)
        Call UpdateRabbitCage(False)
        RCage = True
        MainForm.NewsLabel.Caption = "News: You built a rabbit cage."
    Else
        MainForm.NewsLabel.Caption = "News: You lack the required components."
    End If
End Sub

Sub RabbitMating()
    TRN = RandomInteger(Rabbits + 1, 20)
    If TRN >= 15 Then
        Call UpdateRabbits(1, False)
        MsgBox "One of your rabbits gave birth!"
    End If
End Sub

Sub CraftLiveTrap()
    If RCage = False Then
        MainForm.NewsLabel.Caption = "News: You need to build a rabbit cage first."
        Exit Sub
    End If
    If Sticks >= 10 And Berries >= 3 And Vines >= 5 Then
        Call UpdateStick(-10, False)
        Call UpdateBerry(-3, False)
        Call UpdateLiveTraps(1, False)
        Call UpdateVines(-5, False)
        MainForm.NewsLabel.Caption = "News: You built a live trap."
        LTrapsSet = LTrapsSet + 1
    Else
        MainForm.NewsLabel.Caption = "News: You lack the required components."
    End If
End Sub

Sub CheckLiveTraps()
    Temp1 = LTrapsSet
    For i = 1 To Temp1
        TRN = RandomInteger(1, 10)
        If TRN < 3 Then
            Call UpdateLiveTraps(-1, False)
            MsgBox "One of your live traps broke."
            LTrapsSet = LTrapsSet - 1
        End If
        If TRN >= 7 Then
            Call UpdateRabbits(1, False)
            MsgBox "One of your traps caught a live rabbit! It will need to be reset to catch another."
            LTrapsSet = LTrapsSet - 1
        End If
    Next i
End Sub

Sub MaintainLiveTraps()
    If LTrapsSet = LTraps Then
        MainForm.NewsLabel.Caption = "News: All of your live traps have berries."
        Exit Sub
    End If
    If Berries >= 3 Then
        Call UpdateBerry(-3, False)
        MainForm.NewsLabel.Caption = "News: You put more berries in a trap."
        LTrapsSet = LTrapsSet + 1
        Exit Sub
    Else
        MainForm.NewsLabel.Caption = "News: You don't have enough berries."
        Exit Sub
    End If
End Sub

Sub MaintainBirdTraps()
    If BTrapsSet = BTraps Then
        MainForm.NewsLabel.Caption = "News: All of your bird traps have berries."
        Exit Sub
    End If
    If Berries >= 3 Then
        Call UpdateBerry(-3, False)
        MainForm.NewsLabel.Caption = "News: You put more berries in a trap."
        BTrapsSet = BTrapsSet + 1
        Exit Sub
    Else
        MainForm.NewsLabel.Caption = "News: You don't have enough berries."
        Exit Sub
    End If
End Sub

Sub CheckBirdTraps()
    Temp1 = BTrapsSet
    For i = 1 To Temp1
        TRN = RandomInteger(1, 10)
        If TRN < 3 Then
            Call UpdateBirdTraps(-1, False)
            MsgBox "One of your bird traps broke."
            BTrapsSet = BTrapsSet - 1
        End If
        If TRN >= 7 Then
            Call UpdateFeathers(RandomInteger(2, 10), False)
            Call UpdateRawMeat(1, False)
            MsgBox "One of your traps caught a bird! It will need to be reset to catch another."
            BTrapsSet = BTrapsSet - 1
        End If
    Next i
End Sub

Sub CraftBirdTrap()
    If Sticks >= 10 And Berries >= 3 And Vines >= 5 Then
        Call UpdateStick(-10, False)
        Call UpdateBerry(-3, False)
        Call UpdateBirdTraps(1, False)
        Call UpdateVines(-5, False)
        MainForm.NewsLabel.Caption = "News: You built a bird trap."
        BTrapsSet = BTrapsSet + 1
    Else
        MainForm.NewsLabel.Caption = "News: You lack the required components."
    End If
End Sub

Sub KillRabbit()
    If Rabbits >= 1 Then
        If Knife = True Then Call UpdateFurs(RandomInteger(1, 3), False)
        Call UpdateRawMeat(1, False)
        Call UpdateRabbits(-1, False)
        MainForm.NewsLabel.Caption = "News: You killed a rabbit."
    Else
        MainForm.NewsLabel.Caption = "News: You need a rabbit to kill."
    End If
End Sub

Sub GatherVines()
    FailCount = 0
    TRN = RandomInteger(GatheringLvl, 20)
    Temp1 = 10
    Call UpdateEnergy(-5, False)
    For i = 1 To 10
        RangeS = 20
        RangeE = 22
        RangeS = RangeS - (2 * i)
        RangeE = RangeE - (2 * i)
        If TRN > RangeS And TRN <= RangeE Then
            Call UpdateVines(Temp1 - (i), False)
            MainForm.NewsLabel.Caption = "News: You found " & (Temp1 - (i)) & " vines!"
        Else
            FailCount = FailCount + 1
        End If
    Next i
    If FailCount = 10 Then MainForm.NewsLabel.Caption = "News: You failed to find any vines."
End Sub

Sub FeedRabbits()
    If Berries >= 2 Then
        Call UpdateBerry(-2, False)
        Call UpdateRabbitHunger(5, False)
        MainForm.NewsLabel.Caption = "News: You fed your rabbits some berries."
    Else
        MainForm.NewsLabel.Caption = "News: You have no berries to feed your rabbits."
    End If
End Sub

Sub SaveGame(Slot As Integer)
    ReDim Stats(1 To 88)
    Stats(1) = "Health|" & Health
    Stats(2) = "Hunger|" & Hunger
    Stats(3) = "Thirst|" & Thirst
    Stats(4) = "Energy|" & Energy
    Stats(5) = "MaxHealth|" & MaxHealth
    Stats(6) = "MaxEnergy|" & MaxEnergy
    Stats(7) = "MaxHunger|" & MaxHunger
    Stats(8) = "MaxThirst|" & MaxThirst
    Stats(9) = "DayNum|" & DayNum
    Stats(10) = "Berries|" & Berries
    Stats(11) = "Water|" & Water
    Stats(12) = "Sticks|" & Sticks
    Stats(13) = "Logs|" & Logs
    Stats(14) = "Leaves|" & Leaves
    Stats(15) = "Stones|" & Stones
    Stats(16) = "Inventory|" & Inventory
    Stats(17) = "MaxInventory|" & MaxInventory
    If Spear = True Then
        Stats(18) = "Spear|True"
    Else
        Stats(18) = "Spear|False"
    End If
    If Fire = True Then
        Stats(19) = "Fire|True"
    Else
        Stats(19) = "Fire|False"
    End If
    Stats(20) = "Meat|" & Meat
    Stats(21) = "EnergyPerNigh|" & EnergyPerNight
    If StickShelt = True Then
        Stats(22) = "StickShelt|True"
    Else
        Stats(22) = "StickShelt|False"
    End If
    Stats(23) = "ThingsCooked|" & ThingsCooked
    Stats(24) = "AnimalsHunted|" & AnimalsHunted
    Stats(25) = "HuntingLvl|" & HuntingLvl
    Stats(26) = "CookingLvl|" & CookingLvl
    Stats(27) = "StrengthLvl|" & StrengthLvl
    Stats(28) = "GatheringLvl|" & GatheringLvl
    Stats(29) = "StaminaLvl|" & StaminaLvl
    Stats(30) = "PlayerLvl|" & PlayerLvl
    Stats(31) = "HealthLvl|" & HealthLvl
    Stats(32) = "AllCollected|" & AllCollected
    If Axe = True Then
        Stats(33) = "Axe|True"
    Else
        Stats(33) = "Axe|False"
    End If
    Stats(34) = "CraftingLevel|" & CraftingLevel
    Stats(35) = "HuntThresh|" & HuntThresh
    Stats(36) = "PlayerThresh|" & PlayerThresh
    Stats(37) = "StaminaThresh|" & StaminaThresh
    Stats(38) = "GathThresh|" & GathThresh
    Stats(39) = "StrengthThresh|" & StrengthThresh
    Stats(40) = "CookThresh|" & CookThresh
    Stats(41) = "CraftThresh|" & CraftThresh
    Stats(42) = "ThingsCrafted|" & ThingsCrafted
    Stats(43) = "EnergySpent|" & EnergySpent
    Stats(44) = "HealthGain|" & HealthGain
    If Cabin = True Then
        Stats(45) = "Cabin|True"
    Else
        Stats(45) = "Cabin|False"
    End If
    If Bed = True Then
        Stats(46) = "Bed|True"
    Else
        Stats(46) = "Bed|False"
    End If
    Stats(47) = "Furs|" & Furs
    If Knife = True Then
        Stats(48) = "Knife|True"
    Else
        Stats(48) = "Knife|False"
    End If
    Stats(49) = "RawMeat|" & RawMeat
    Stats(50) = "AllLevels|" & AllLevels
    Stats(51) = "ActionsCompleted|" & ActionsCompleted
    Stats(52) = "Traps|" & Traps
    Stats(53) = "FireStartDay|" & FireStartDay
    Stats(54) = "SetTraps|" & SetTraps
    If CWolf = True Then
        Stats(55) = "CWolf|True"
    Else
        Stats(55) = "CWolf|False"
    End If
    Stats(56) = "CWolfLvl|" & CWolfLvl
    Stats(57) = "CWolfHealth|" & CWolfHealth
    Stats(58) = "CWolfHunger|" & CWolfHunger
    Stats(59) = "CWolfBond|" & CWolfBond
    Stats(60) = "CWolfBondLvl|" & CWolfBondLvl
    Stats(61) = "CWolfMaxHealth|" & CWolfMaxHealth
    Stats(62) = "CWolfMaxHunger|" & CWolfMaxHunger
    Stats(63) = "CWolfXP|" & CWolfXP
    Stats(64) = "CWolfThresh|" & CWolfThresh
    Stats(65) = "CWolfBondThresh|" & CWolfBondThresh
    If CWolfDailyAction = True Then
        Stats(66) = "CWolfDailyAction|True"
    Else
        Stats(66) = "CWolfDailyAction|False"
    End If
    Stats(67) = "LastCWolfBondThresh|" & LastCWolfBondThresh
    If WolfPrompt1 = True Then
        Stats(68) = "WolfPrompt1|True"
    Else
        Stats(68) = "WolfPrompt1|False"
    End If
    Stats(69) = "WolfBond|" & WolfBond
    If WolfPrompt2 = True Then
        Stats(70) = "WolfPrompt2|True"
    Else
        Stats(70) = "WolfPrompt2|True"
    End If
    If WolfPrompt3 = True Then
        Stats(71) = "WolfPrompt3|True"
    Else
        Stats(71) = "WolfPrompt3|True"
    End If
    Stats(72) = "CWolfName|" & CWolfName
    Stats(73) = "Arrows|" & Arrows
    Stats(74) = "Vines|" & Vines
    If Bow = True Then
        Stats(75) = "Bow|True"
    Else
        Stats(75) = "Bow|False"
    End If
    Stats(76) = "LTraps|" & LTraps
    Stats(77) = "LTrapsSet|" & LTrapsSet
    Stats(78) = "BTraps|" & BTraps
    Stats(79) = "BTrapsSet|" & BTrapsSet
    If RCage = True Then
        Stats(80) = "RCage|True"
    Else
        Stats(80) = "RCage|False"
    End If
    Stats(81) = "Rabbits|" & Rabbits
    Stats(82) = "Feathers|" & Feathers
    Stats(83) = "MaxRabbits|" & MaxRabbits
    Stats(84) = "RabbitHunger|" & RabbitHunger
    Stats(85) = "MaxRabbitHunger|" & MaxRabbitHunger
    Stats(86) = "AutoSave|" & AutoSave
    Stats(87) = "SelectedSlot|" & SelectedSlot
    Stats(88) = "HealthThresh|" & HealthThresh
    If Slot = 1 Then Call SaveTextFile(PathToDesktop & "\SaveFile1.txt", 1, 88, Stats)
    If Slot = 2 Then Call SaveTextFile(PathToDesktop & "\SaveFile2.txt", 1, 88, Stats)
    If Slot = 3 Then Call SaveTextFile(PathToDesktop & "\SaveFile3.txt", 1, 88, Stats)
    MsgBox "Game saved."
End Sub

Sub LoadGame(Slot As Integer)
    LoadedGame = True
    If Slot = 1 Then Call LoadTextFile(PathToDesktop & "\SaveFile1.txt", Stats)
    If Slot = 2 Then Call LoadTextFile(PathToDesktop & "\SaveFile2.txt", Stats)
    If Slot = 3 Then Call LoadTextFile(PathToDesktop & "\SaveFile3.txt", Stats)
    Health = Right(Stats(1), Len(Stats(1)) - InStr(1, Stats(1), "|"))
    Hunger = Right(Stats(2), Len(Stats(2)) - InStr(1, Stats(2), "|"))
    Thirst = Right(Stats(3), Len(Stats(3)) - InStr(1, Stats(3), "|"))
    Energy = Right(Stats(4), Len(Stats(4)) - InStr(1, Stats(4), "|"))
    MaxHealth = Right(Stats(5), Len(Stats(5)) - InStr(1, Stats(5), "|"))
    MaxEnergy = Right(Stats(6), Len(Stats(6)) - InStr(1, Stats(6), "|"))
    MaxHunger = Right(Stats(7), Len(Stats(7)) - InStr(1, Stats(7), "|"))
    MaxThirst = Right(Stats(8), Len(Stats(8)) - InStr(1, Stats(8), "|"))
    DayNum = Right(Stats(9), Len(Stats(9)) - InStr(1, Stats(9), "|"))
    Berries = Right(Stats(10), Len(Stats(10)) - InStr(1, Stats(10), "|"))
    Water = Right(Stats(11), Len(Stats(11)) - InStr(1, Stats(11), "|"))
    Sticks = Right(Stats(12), Len(Stats(12)) - InStr(1, Stats(12), "|"))
    Logs = Right(Stats(13), Len(Stats(13)) - InStr(1, Stats(13), "|"))
    Leaves = Right(Stats(14), Len(Stats(14)) - InStr(1, Stats(14), "|"))
    Stones = Right(Stats(15), Len(Stats(15)) - InStr(1, Stats(15), "|"))
    Inventory = Right(Stats(16), Len(Stats(16)) - InStr(1, Stats(16), "|"))
    MaxInventory = Right(Stats(17), Len(Stats(17)) - InStr(1, Stats(17), "|"))
    If Right(Stats(18), Len(Stats(18)) - InStr(1, Stats(18), "|")) = "True" Then
        Spear = True
    Else
        Spear = False
    End If
    If Right(Stats(19), Len(Stats(19)) - InStr(1, Stats(19), "|")) = "True" Then
        Fire = True
    Else
        Fire = False
    End If
    Meat = Right(Stats(20), Len(Stats(20)) - InStr(1, Stats(20), "|"))
    EnergyPerNight = Right(Stats(21), Len(Stats(21)) - InStr(1, Stats(21), "|"))
    If Right(Stats(22), Len(Stats(22)) - InStr(1, Stats(22), "|")) Then
        StickShelt = True
    Else
        StickShelt = False
    End If
    ThingsCooked = Right(Stats(23), Len(Stats(23)) - InStr(1, Stats(23), "|"))
    AnimalsHunted = Right(Stats(24), Len(Stats(24)) - InStr(1, Stats(24), "|"))
    HuntingLvl = Right(Stats(25), Len(Stats(25)) - InStr(1, Stats(25), "|"))
    CookingLvl = Right(Stats(26), Len(Stats(26)) - InStr(1, Stats(26), "|"))
    StrengthLvl = Right(Stats(27), Len(Stats(27)) - InStr(1, Stats(27), "|"))
    GatheringLvl = Right(Stats(28), Len(Stats(28)) - InStr(1, Stats(28), "|"))
    StaminaLvl = Right(Stats(29), Len(Stats(29)) - InStr(1, Stats(29), "|"))
    PlayerLvl = Right(Stats(30), Len(Stats(30)) - InStr(1, Stats(30), "|"))
    HealthLvl = Right(Stats(31), Len(Stats(31)) - InStr(1, Stats(31), "|"))
    AllCollected = Right(Stats(32), Len(Stats(32)) - InStr(1, Stats(32), "|"))
    If Right(Stats(33), Len(Stats(33)) - InStr(1, Stats(33), "|")) = "True" Then
        Axe = True
    Else
        Axe = False
    End If
    CraftingLevel = Right(Stats(34), Len(Stats(34)) - InStr(1, Stats(34), "|"))
    HuntThresh = Right(Stats(35), Len(Stats(35)) - InStr(1, Stats(35), "|"))
    PlayerThresh = Right(Stats(36), Len(Stats(36)) - InStr(1, Stats(36), "|"))
    StaminaThresh = Right(Stats(37), Len(Stats(37)) - InStr(1, Stats(37), "|"))
    GathThresh = Right(Stats(38), Len(Stats(38)) - InStr(1, Stats(38), "|"))
    StrengthThresh = Right(Stats(39), Len(Stats(39)) - InStr(1, Stats(39), "|"))
    CookThresh = Right(Stats(40), Len(Stats(40)) - InStr(1, Stats(40), "|"))
    CraftThresh = Right(Stats(41), Len(Stats(41)) - InStr(1, Stats(41), "|"))
    ThingsCrafted = Right(Stats(42), Len(Stats(42)) - InStr(1, Stats(42), "|"))
    EnergySpent = Right(Stats(43), Len(Stats(43)) - InStr(1, Stats(43), "|"))
    HealthGain = Right(Stats(44), Len(Stats(44)) - InStr(1, Stats(44), "|"))
    If Right(Stats(45), Len(Stats(45)) - InStr(1, Stats(45), "|")) = "True" Then
        Cabin = True
    Else
        Cabin = False
    End If
    If Right(Stats(46), Len(Stats(46)) - InStr(1, Stats(46), "|")) = "True" Then
        Bed = True
    Else
        Bed = False
    End If
    Furs = Right(Stats(47), Len(Stats(47)) - InStr(1, Stats(47), "|"))
    If Right(Stats(48), Len(Stats(48)) - InStr(1, Stats(48), "|")) = "True" Then
        Knife = True
    Else
        Knife = False
    End If
    RawMeat = Right(Stats(49), Len(Stats(49)) - InStr(1, Stats(49), "|"))
    AllLevels = Right(Stats(50), Len(Stats(50)) - InStr(1, Stats(50), "|"))
    ActionsCompleted = Right(Stats(51), Len(Stats(51)) - InStr(1, Stats(51), "|"))
    Traps = Right(Stats(52), Len(Stats(52)) - InStr(1, Stats(52), "|"))
    FireStartDay = Right(Stats(53), Len(Stats(53)) - InStr(1, Stats(53), "|"))
    SetTraps = Right(Stats(54), Len(Stats(54)) - InStr(1, Stats(54), "|"))
    If Right(Stats(55), Len(Stats(55)) - InStr(1, Stats(55), "|")) = "True" Then
        CWolf = True
    Else
        CWolf = False
    End If
    CWolfLvl = Right(Stats(56), Len(Stats(56)) - InStr(1, Stats(56), "|"))
    CWolfHealth = Right(Stats(57), Len(Stats(57)) - InStr(1, Stats(57), "|"))
    CWolfHunger = Right(Stats(58), Len(Stats(58)) - InStr(1, Stats(58), "|"))
    CWolfBond = Right(Stats(59), Len(Stats(59)) - InStr(1, Stats(59), "|"))
    CWolfBondLvl = Right(Stats(60), Len(Stats(60)) - InStr(1, Stats(60), "|"))
    CWolfMaxHealth = Right(Stats(61), Len(Stats(61)) - InStr(1, Stats(61), "|"))
    CWolfMaxHunger = Right(Stats(62), Len(Stats(62)) - InStr(1, Stats(62), "|"))
    CWolfXP = Right(Stats(63), Len(Stats(63)) - InStr(1, Stats(63), "|"))
    CWolfThresh = Right(Stats(64), Len(Stats(64)) - InStr(1, Stats(64), "|"))
    CWolfBondThresh = Right(Stats(65), Len(Stats(65)) - InStr(1, Stats(65), "|"))
    If Right(Stats(66), Len(Stats(66)) - InStr(1, Stats(66), "|")) = "True" Then
        CWolfDailyAction = True
    Else
        CWolfDailyAction = False
    End If
    LastCWolfBondThresh = Right(Stats(67), Len(Stats(67)) - InStr(1, Stats(67), "|"))
    If Right(Stats(68), Len(Stats(68)) - InStr(1, Stats(68), "|")) = "True" Then
        WolfPrompt1 = True
    Else
        WolfPrompt1 = False
    End If
    WolfBond = Right(Stats(69), Len(Stats(69)) - InStr(1, Stats(69), "|"))
    If Right(Stats(70), Len(Stats(70)) - InStr(1, Stats(70), "|")) = "True" Then
        WolfPrompt2 = True
    Else
        WolfPrompt2 = False
    End If
    If Right(Stats(71), Len(Stats(71)) - InStr(1, Stats(71), "|")) = "True" Then
        WolfPrompt3 = True
    Else
        WolfPrompt3 = False
    End If
    CWolfName = Right(Stats(72), Len(Stats(72)) - InStr(1, Stats(72), "|"))
    Arrows = Right(Stats(73), Len(Stats(73)) - InStr(1, Stats(73), "|"))
    Vines = Right(Stats(74), Len(Stats(74)) - InStr(1, Stats(74), "|"))
    If Right(Stats(75), Len(Stats(75)) - InStr(1, Stats(75), "|")) = "True" Then
        Bow = True
    Else
        Bow = False
    End If
    LTraps = Right(Stats(76), Len(Stats(76)) - InStr(1, Stats(76), "|"))
    LTrapsSet = Right(Stats(77), Len(Stats(77)) - InStr(1, Stats(77), "|"))
    BTraps = Right(Stats(78), Len(Stats(78)) - InStr(1, Stats(78), "|"))
    BTrapsSet = Right(Stats(79), Len(Stats(79)) - InStr(1, Stats(79), "|"))
    If Right(Stats(80), Len(Stats(80)) - InStr(1, Stats(80), "|")) = "True" Then
        RCage = True
    Else
        RCage = False
    End If
    Rabbits = Right(Stats(81), Len(Stats(81)) - InStr(1, Stats(81), "|"))
    Feathers = Right(Stats(82), Len(Stats(82)) - InStr(1, Stats(82), "|"))
    MaxRabbits = Right(Stats(83), Len(Stats(83)) - InStr(1, Stats(83), "|"))
    RabbitHunger = Right(Stats(84), Len(Stats(84)) - InStr(1, Stats(84), "|"))
    MaxRabbitHunger = Right(Stats(85), Len(Stats(85)) - InStr(1, Stats(85), "|"))
    AutoSave = Right(Stats(86), Len(Stats(86)) - InStr(1, Stats(86), "|"))
    SelectedSlot = Right(Stats(87), Len(Stats(87)) - InStr(1, Stats(87), "|"))
    HealthThresh = Right(Stats(88), Len(Stats(88)) - InStr(1, Stats(88), "|"))
    Call UpdateAll
    MsgBox "Game loaded."
    Call UpdateOpenFrame("GameFrame")
    LoadedGame = False
End Sub

Sub HarvestCrops()
    For i = 1 To MaxCrops
        If Crop(i) = "Corn" Then
            If CropGrowth(i) = CornGP Then
                Call UpdateCorn(1, False)
                Call UpdateCornSeeds(1, False)
                If RandomInteger(1, 3) = 1 Then Call UpdateCornSeeds(1, False)
            End If
        ElseIf Crop(i) = "Berry" Then
            If CropGrowth(i) = BerryGP Then
                Call UpdateBerry(1, False)
                If RandomInteger(1, 3) = 1 Then Call UpdateBerry(1, False)
            End If
        End If
    Next i
End Sub
