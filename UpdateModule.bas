Attribute VB_Name = "Updates"

Sub UpdateHealth(Amount As Integer, Reset As Boolean) 'Increases Health by Amount and updates HealthLabel
    If Amount < 0 Then
        If Amount < -Health Then Amount = -Health
    Else
        If Amount > MaxHealth - Health Then Amount = MaxHealth - Health
    End If
    If Reset = False Then
        Health = Health + Amount
        If Amount > 0 Then Call UpdateHealthGain(Amount, False)
    Else
        MaxHealth = 100
        Health = MaxHealth
    End If
    MainForm.HealthLabel.Caption = "Health: " & Health & "/" & MaxHealth
    If Health = 0 Then
        MsgBox ("You died!")
        GameInProgress = False
        Call UpdateOpenFrame("MainMenu")
        Exit Sub
    End If
End Sub

Sub UpdateBerry(Amount As Integer, Reset As Boolean) 'Increases berries by amount and Updates BerryLabel in inventory
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        If Amount > 0 Then BerriesCollected = BerriesCollected + Amount
        If Amount > 0 Then Call UpdateAllCollected(Amount, False)
        Berries = Berries + Amount
        Call UpdateInventory(Amount, False)
    Else
        Berries = 0
    End If
    MainForm.BerryLabel.Caption = Berries & " Berries"
End Sub

Sub UpdateFurs(Amount As Integer, Reset As Boolean) 'Increases furs by amount and Updates BerryLabel in inventory
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        Furs = Furs + Amount
        Call UpdateInventory(Amount, False)
    Else
        Furs = 0
    End If
    MainForm.FurLabel.Caption = Furs & " Furs"
End Sub

Sub UpdateLeaf(Amount As Integer, Reset As Boolean) 'Increases Leaves by amount and Updates LeafLabel in inventory
    If Reset = False Then
        If Amount > 0 Then LeavesCollected = LeavesCollected + Amount
        'If Amount > 0 Then Call UpdateAllCollected(Amount, False)
        Leaves = Leaves + Amount
        'Call UpdateInventory(Amount, False)
    Else
        Leaves = 0
    End If
    MainForm.LeafLabel.Caption = Leaves & " Leaves"
End Sub

Sub UpdateLog(Amount As Integer, Reset As Boolean) 'Increases logs by amount and Updates LogLabel in inventory
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        If Amount > 0 Then LogsCollected = LogsCollected + Amount
        If Amount > 0 Then Call UpdateAllCollected(Amount, False)
        Logs = Logs + Amount
        Call UpdateInventory(Amount * 5, False)
    Else
        Logs = 0
    End If
    MainForm.LogLabel.Caption = Logs & " Logs"
End Sub

Sub UpdateStick(Amount As Integer, Reset As Boolean) 'Increases sticks by amound and Updates StickLabel in inventory
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        If Amount > 0 Then SticksCollected = SticksCollected + Amount
        If Amount > 0 Then Call UpdateAllCollected(Amount, False)
        Sticks = Sticks + Amount
        Call UpdateInventory(Amount, False)
    Else
        Sticks = 0
    End If
    MainForm.StickLabel.Caption = Sticks & " Sticks"
End Sub

Sub UpdateStone(Amount As Integer, Reset As Boolean) 'Increases stone by amount and Updates StickLabel in inventory
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        If Amount > 0 Then StonesCollected = StonesCollected + Amount
        If Amount > 0 Then Call UpdateAllCollected(Amount, False)
        Stones = Stones + Amount
        Call UpdateInventory(Amount, False)
    Else
        Stones = 0
    End If
    MainForm.StoneLabel.Caption = Stones & " Stones"
End Sub
Sub UpdateWater(Amount As Integer, Reset As Boolean) 'Increases water by amount and Updates WaterLabel in inventory
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        If Amount > 0 Then WaterCollected = WaterCollected + Amount
        If Amount > 0 Then Call UpdateAllCollected(Amount, False)
        Water = Water + Amount
        Call UpdateInventory(Amount, False)
    Else
        Water = 0
    End If
    MainForm.WaterLabel.Caption = Water & " Cups of Water"
End Sub

Sub UpdateInventory(Amount As Integer, Reset As Boolean) 'Increases Inventory by Amount and updates Inventory frame label
    If Reset = False Then
        Inventory = Inventory + Amount
    Else
        Inventory = 0
        MaxInventory = 50
    End If
    MainForm.InventoryFrame.Caption = "Inventory: " & Inventory & "/" & MaxInventory
    If MaxInventory - Inventory = 0 Then MainForm.NewsLabel.Caption = "News: Your inventory is full."
End Sub

Sub UpdateHunger(Amount As Integer, Reset As Boolean) 'Increases Hunger by Amount and updates HungerLabel
    If Amount < 0 Then
        If Amount < -Hunger Then Amount = -Hunger
    Else
        If Amount > MaxHunger - Hunger Then Amount = MaxHunger - Hunger
    End If
    If Reset = False Then
        Hunger = Hunger + Amount
    Else
        MaxHunger = 25
        Hunger = MaxHunger
    End If
    MainForm.HungerLabel.Caption = "Hunger: " & Hunger & "/" & MaxHunger
    If Hunger = 0 Then
        Call UpdateHealth(-10, False)
    End If
End Sub

Sub UpdateEnergy(Amount As Integer, Reset As Boolean) 'Increases Energy by Amount and updates EnergyLabel
    If Amount < 0 Then
        If Amount < -Energy Then Amount = -Energy
    Else
        If Amount > MaxEnergy - Energy Then Amount = MaxEnergy - Energy
    End If
    If Reset = False Then
        Energy = Energy + Amount
        If Amount < 0 Then Call UpdateEnergySpent(Amount, False)
    Else
        MaxEnergy = 25
        Energy = MaxEnergy
    End If
    MainForm.EnergyLabel.Caption = "Energy: " & Energy & "/" & MaxEnergy
End Sub

Sub UpdateThirst(Amount As Integer, Reset As Boolean) 'Increases Thirst by Amount and updates ThirstLabel
    If Amount < 0 Then
        If Amount < -Thirst Then Amount = -Thirst
    Else
        If Amount > MaxThirst - Thirst Then Amount = MaxThirst - Thirst
    End If
    If Reset = False Then
        Thirst = Thirst + Amount
    Else
        MaxThirst = 25
        Thirst = MaxThirst
    End If
    MainForm.ThirstLabel.Caption = "Thirst: " & Thirst & "/" & MaxThirst
    If Thirst = 0 Then
        Call UpdateHealth(-10, False)
    End If
End Sub

Sub UpdateMeat(Amount As Integer, Reset As Boolean) 'Increases meat by Amount and updates MeatLabel
    If Reset = False Then
        Meat = Meat + Amount
    Else
        Meat = 0
    End If
    MainForm.MeatLabel.Caption = Meat & " Meat"
End Sub

Sub UpdateRawMeat(Amount As Integer, Reset As Boolean) 'Increases raw meat amount and updates the uncooked rabbit label
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        RawMeat = RawMeat + Amount
        'If Amount > 0 Then Call UpdateAnimalsHunted(Amount, False)
        Call UpdateInventory(Amount, False)
    Else
        RawMeat = 0
    End If
    MainForm.RawMeatLabel.Caption = RawMeat & " Raw Meat"
End Sub

Sub UpdateDay(Reset As Boolean) 'Increases the day number by 1 or Resets day number to 1
    If Reset = True Then
        DayNum = 1
    Else
        DayNum = DayNum + 1
    End If
    MainForm.DayLabel.Caption = "Day #" & DayNum
End Sub

Sub UpdateEnergyPerNight() 'Changes Energypernight
    EnergyPerNight = 10
    If StickShelt = True Then EnergyPerNight = EnergyPerNight + 5
    If Cabin = True Then EnergyPerNight = EnergyPerNight + 5
    If Bed = True Then EnergyPerNight = EnergyPerNight + 5
End Sub

Sub UpdateThingsCooked(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        ThingsCooked = ThingsCooked + Amount
        If ThingsCooked = CookThresh Then
            CookingLvl = CookingLvl + 1
            CookThresh = CookThresh + (5 * CookingLvl)
            MainForm.CookLvlLabel.Caption = "Cooking: " & CookingLvl
            MsgBox "News: Your cooking level increased!"
            Call UpdateAllLevels(1, False)
        End If
    Else
        ThingsCooked = 0
        CookingLvl = 1
        CookThresh = 5
        MainForm.CookLvlLabel.Caption = "Cooking: " & CookingLvl
        Call UpdateAllLevels(1, False)
    End If
End Sub

Sub UpdateEnergySpent(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        EnergySpent = EnergySpent + Amount
        If EnergySpent <= -StaminaThresh Then
            StaminaLvl = StaminaLvl + 1
            StaminaThresh = StaminaThresh + (40 * StaminaLvl)
            MainForm.StamLvlLabel.Caption = "Stamina: " & StaminaLvl
            MsgBox "News: Your stamina level increased!"
            MaxEnergy = MaxEnergy + (5)
            Call UpdateAllLevels(1, False)
        End If
    Else
        EnergySpent = 0
        StaminaLvl = 1
        StaminaThresh = 30
        MainForm.StamLvlLabel.Caption = "Stamina: " & StaminaLvl
        Call UpdateAllLevels(1, False)
    End If
End Sub

Sub UpdateAnimalsHunted(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        AnimalsHunted = AnimalsHunted + Amount
        If AnimalsHunted >= HuntThresh Then
            HuntingLvl = HuntingLvl + 1
            HuntThresh = HuntThresh + (2 * StaminaLvl)
            MainForm.HuntLvlLabel.Caption = "Hunting: " & HuntingLvl
            MsgBox "News: Your hunting level increased!"
            Call UpdateAllLevels(1, False)
        End If
    Else
        AnimalsHunted = 0
        HuntingLvl = 1
        HuntThresh = 3
        MainForm.HuntLvlLabel.Caption = "Hunting: " & HuntingLvl
        Call UpdateAllLevels(1, False)
    End If
End Sub

Sub UpdateAllCollected(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        AllCollected = AllCollected + Amount
        If AllCollected >= GathThresh Then
            GatheringLvl = GatheringLvl + 1
            GathThresh = GathThresh + (20 * GatheringLvl)
            MainForm.GathLvlLabel.Caption = "Gathering: " & GatheringLvl
            MsgBox "News: Your gathering level increased!"
            Call UpdateAllLevels(1, False)
        End If
    Else
        AllCollected = 0
        GatheringLvl = 1
        GathThresh = 30
        MainForm.GathLvlLabel.Caption = "Gathering: " & GatheringLvl
        Call UpdateAllLevels(1, False)
    End If
End Sub

Sub UpdateThingsCrafted(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        ThingsCrafted = ThingsCrafted + Amount
        If ThingsCrafted >= CraftThresh Then
            CraftingLevel = CraftingLevel + 1
            CraftThresh = CraftThresh + (2 * CraftingLevel)
            MainForm.CraftLvlLabel.Caption = "Crafting: " & CraftingLevel
            MsgBox "News: Your crafting level increased!"
            Call UpdateAllLevels(1, False)
        End If
    Else
        ThingsCrafted = 0
        CraftingLevel = 1
        CraftThresh = 1
        MainForm.CraftLvlLabel.Caption = "Crafting: " & CraftingLevel
        Call UpdateAllLevels(1, False)
    End If
End Sub

Sub UpdateHealthGain(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        HealthGain = HealthGain + Amount
        If HealthGain >= HealthThresh Then
            HealthLvl = HealthLvl + 1
            HealthThresh = HealthThresh + (5 * HealthLvl)
            MainForm.HealLvlLabel.Caption = "Health: " & HealthLvl
            MsgBox "News: Your health level increased!"
            MaxHealth = MaxHealth + (5)
            Call UpdateAllLevels(1, False)
        End If
    Else
        HealthGain = 0
        HealthLvl = 1
        HealthThresh = 10
        MainForm.HealLvlLabel.Caption = "Health: " & HealthLvl
        Call UpdateAllLevels(1, False)
    End If
End Sub

Sub UpdateConsoleOpen()
    If ConsoleOpen = False Then
        ConsoleOpen = True
        MainForm.GameFrame.Height = 6015
        MainForm.Console.Visible = True
    Else
        ConsoleOpen = False
        MainForm.GameFrame.Height = 5535
        MainForm.Console.Visible = False
    End If
End Sub

Sub UpdateAllLevels(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        AllLevels = AllLevels + Amount
        If AllLevels >= PlayerThresh Then
            PlayerLvl = PlayerLvl + 1
            PlayerThresh = PlayerThresh + 8
            MainForm.LevelsFrame.Caption = "Level: " & PlayerLvl
            MsgBox "News: Your overall level increased!"
            AllLevels = AllLevels + 1
        End If
    Else
        AllLevels = 0
        PlayerLvl = 1
        PlayerThresh = 8
        MainForm.LevelsFrame.Caption = "Level: " & PlayerLvl
    End If
End Sub

Sub UpdateActionsCompleted(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        ActionsCompleted = ActionsCompleted + Amount
        If ActionsCompleted >= StrengthThresh Then
            StrengthLvl = StrengthLvl + 1
            StrengthThresh = StrengthThresh + (10 * StrengthLvl)
            MainForm.StrenLvlLabel.Caption = "Strength: " & StrengthLvl
            MsgBox "News: Your strength level increased!"
            Call UpdateAllLevels(1, False)
            MaxInventory = MaxInventory + 25
            Call UpdateInventory(0, False)
        End If
    Else
        ActionsCompleted = 0
        StrengthLvl = 1
        StrengthThresh = 10
        MainForm.StrenLvlLabel.Caption = "Strength: " & StrengthLvl
        Call UpdateAllLevels(1, False)
    End If
End Sub

Sub UpdateTraps(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        Traps = Traps + Amount
    Else
        Traps = 0
    End If
    MainForm.TrapLabel.Caption = Traps & " Traps"
End Sub

Sub UpdateCWolf(Reset As Boolean)
    If Reset = False Then
        CWolf = True
        If CWolfName = "" Then CWolfName = "Worf the wolf"
        MainForm.WolfLabel.Caption = CWolfName
        MainForm.WolfLabel.Visible = True
        If LoadedGame = False Then MsgBox "The wolf is now your companion! It will complete tasks for you in exchange for occasional scraps. If you don't feed it though, it will run away or die."
    Else
        CWolf = False
        MainForm.WolfLabel.Caption = "Wolf"
        MainForm.WolfLabel.Visible = False
        MainForm.BackButton2.Visible = False
        MainForm.WolfCommandFrame.Visible = False
        MainForm.CompActFrame.Visible = True
        MainForm.WolfInteractFrame.Visible = False
        MainForm.CompActFrame.Caption = "Select a Companion"
        MainForm.CompStatFrame.Caption = "Select a Companion"
        Call UpdateCWolfBond(0, True)
        Call UpdateCWolfBond(2, False)
        Call UpdateCWolfHunger(0, True)
        Call UpdateCWolfHealth(0, True)
        Call UpdateCWolfXP(0, True)
        WolfPrompt1 = False
        WolfPrompt2 = True
        WolfPrompt3 = True
    End If
End Sub

Sub UpdateCWolfXP(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        CWolfXP = CWolfXP + Amount
        If CWolfXP >= CWolfThresh Then
            Call UpdateCWolfLvl(1, False)
            CWolfThresh = CWolfThresh + (CWolfLvl * 15)
        End If
    Else
        CWolfXP = 0
        Call UpdateCWolfLvl(0, True)
        CWolfThresh = 10
    End If
End Sub

Sub UpdateCWolfHealth(Amount As Integer, Reset As Boolean)
    If Amount < 0 Then
        If Amount < -CWolfHealth Then Amount = -CWolfHealth
    Else
        If Amount > CWolfMaxHealth - CWolfHealth Then Amount = CWolfMaxHealth - CWolfHealth
    End If
    If Reset = False Then
        CWolfHealth = CWolfHealth + Amount
    Else
        CWolfMaxHealth = 100
        CWolfHealth = CWolfMaxHealth
    End If
    If MainForm.CompStatFrame.Caption = "Wolf - Stats" Then MainForm.CompHealthLabel.Caption = "Health: " & CWolfHealth & "/" & CWolfMaxHealth
    If CWolfHealth = 0 And CWolf = True Then
        MsgBox ("Your wolf died!")
    End If
End Sub

Sub UpdateCWolfHunger(Amount As Integer, Reset As Boolean) 'Increases Hunger by Amount and updates HungerLabel
    If Amount < 0 Then
        If Amount < -CWolfHunger Then Amount = -CWolfHunger
    Else
        If Amount > CWolfMaxHunger - CWolfHunger Then Amount = CWolfMaxHunger - CWolfHunger
    End If
    If Reset = False Then
        CWolfHunger = CWolfHunger + Amount
    Else
        CWolfMaxHunger = 25
        CWolfHunger = CWolfMaxHunger
    End If
    If MainForm.CompStatFrame.Caption = "Wolf - Stats" Then MainForm.CompHungerLabel.Caption = "Hunger: " & CWolfHunger & "/" & CWolfMaxHunger
    If CWolfHunger = 0 Then
        Call UpdateCWolfHealth(-10, False)
        Call UpdateCWolfBond(-10, False)
    End If
End Sub

Sub UpdateCWolfBond(Amount As Integer, Reset As Boolean)
    If Amount < 0 Then
        If Amount < -CWolfBond Then Amount = -CWolfBond
    Else
        'If Amount > CWolfMaxbond - CWolfBond Then Amount = CWolfMaxHunger - CWolfHunger
    End If
    If Reset = False Then
        CWolfBond = CWolfBond + Amount
    Else
        CWolfBond = 0
        CWolfBondThresh = 10
        Call UpdateCWolfBondLvl(0, True)
    End If
    If CWolfBond >= CWolfBondThresh Then
        Call UpdateCWolfBondLvl(1, False)
        LastCWolfBondThresh = CWolfBondThresh
        CWolfBondThresh = CWolfBondThresh + (CWolfBondLvl * 15)
    ElseIf CWolfBond < LastCWolfBondThresh And CWolfBondLvl <> 1 Then
        Call UpdateCWolfBondLvl(-1, False)
        CWolfBondThresh = CWolfBondThresh - ((CWolfBondLvl - 1) * 15)
    End If
    If CWolfBond = 0 And Reset = False Then
        If CWolf = True Then Call UpdateCWolf(True)
        If CWolf = True Then MsgBox "Your wolf ran away!"
    End If
End Sub

Sub UpdateCWolfLvl(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        CWolfLvl = CWolfLvl + Amount
    Else
        CWolfLvl = 1
    End If
    If MainForm.CompStatFrame.Caption = "Wolf - Stats" Then MainForm.CompLvlLabel.Caption = "Level: " & CWolfLvl
End Sub

Sub UpdateCWolfBondLvl(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        CWolfBondLvl = CWolfBondLvl + Amount
    Else
        CWolfBondLvl = 1
    End If
    If MainForm.CompStatFrame.Caption = "Wolf - Stats" Then MainForm.CompBondLabel.Caption = "Bond Level: " & CWolfBondLvl
End Sub

Sub UpdateWolfBond(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        WolfBond = WolfBond + Amount
        If WolfBond = 2 Then
            UpdateCWolf (False)
        End If
    Else
        WolfBond = 0
    End If
End Sub

Sub UpdateArrows(Amount As Integer, Reset As Boolean)
    'If Amount < 0 Then
        'If Amount < -Inventory Then Amount = -Inventory
    'Else
        'If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    'End If
    If Reset = False Then
        Arrows = Arrows + Amount
        'Call UpdateInventory(Amount, False)
    Else
        Arrows = 0
    End If
    MainForm.ArrowLabel.Caption = Arrows & " Arrows"
End Sub

Sub UpdateVines(Amount As Integer, Reset As Boolean)
    If Amount < 0 Then
        If Amount < -Inventory Then Amount = -Inventory
    Else
        If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    End If
    If Reset = False Then
        Vines = Vines + Amount
        Call UpdateInventory(Amount, False)
    Else
        Vines = 0
    End If
    MainForm.VineLabel.Caption = Vines & " Vines"
End Sub

Sub UpdateFeathers(Amount As Integer, Reset As Boolean)
    'If Amount < 0 Then
        'If Amount < -Inventory Then Amount = -Inventory
    'Else
        'If Amount > MaxInventory - Inventory Then Amount = MaxInventory - Inventory
    'End If
    If Reset = False Then
        Feathers = Feathers + Amount
        'Call UpdateInventory(Amount, False)
    Else
        Feathers = 0
    End If
    MainForm.FeatherLabel.Caption = Feathers & " Feathers"
End Sub

Sub UpdateRabbits(Amount As Integer, Reset As Boolean)
    If Amount < 0 Then
        If Amount < -Rabbits Then Amount = -Rabbits
    Else
        If Amount > MaxRabbits - Rabbits Then Amount = MaxRabbits - Rabbits
    End If
    If Reset = False Then
        Rabbits = Rabbits + Amount
    Else
        Rabbits = 0
        MaxRabbits = 10
    End If
    MainForm.RabbitLabel.Caption = Rabbits & " Rabbits"
End Sub

Sub UpdateRabbitHunger(Amount As Integer, Reset As Boolean) 'Increases Hunger by Amount and updates HungerLabel
    If Amount < 0 Then
        If Amount < -RabbitHunger Then Amount = -RabbitHunger
    Else
        If Amount > MaxRabbitHunger - RabbitHunger Then Amount = MaxRabbitHunger - RabbitHunger
    End If
    If Reset = False Then
        RabbitHunger = RabbitHunger + Amount
    Else
        MaxRabbitHunger = 25
        RabbitHunger = MaxRabbitHunger
    End If
    If MainForm.AnimalStats.Caption = "Rabbit - Stats" Then MainForm.AnimalHungerLabel = "Hunger: " & RabbitHunger & "/" & MaxRabbitHunger
    If RabbitHunger = 0 Then
        Call UpdateRabbits(-1, False)
        MsgBox "One of your rabbits died!"
    End If
End Sub

Sub UpdateBirdTraps(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        BTraps = BTraps + Amount
    Else
        BTraps = 0
    End If
    MainForm.BTrapLabel.Caption = BTraps & " Bird Traps"
End Sub

Sub UpdateLiveTraps(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        LTraps = LTraps + Amount
    Else
        LTraps = 0
    End If
    MainForm.LTrapLabel.Caption = LTraps & " Live Traps"
End Sub

Sub UpdateBow(Reset As Boolean)
    If Reset = False Then
        MainForm.BowLabel.Caption = "Bow"
    Else
        Bow = False
        MainForm.BowLabel.Caption = "No Bow"
    End If
End Sub

Sub UpdateRabbitCage(Reset As Boolean)
    If Reset = False Then
        MainForm.RabbitLabel.Visible = True
    Else
        RCage = False
        MainForm.RabbitLabel.Visible = False
    End If
End Sub

Sub UpdateOpenFrame(Frame As String)
    Select Case Frame
        Case "MainMenu"
            MainForm.Width = MainForm.MainMenuFrame.Width + 100
            MainForm.Height = MainForm.MainMenuFrame.Height + 450
            MainForm.GameFrame.Visible = False
            MainForm.MainMenuFrame.Visible = True
            MainForm.OptionFrame.Visible = False
            MainForm.SaveGameFrame.Visible = False
            MainForm.LoadGameFrame.Visible = False
            If GameInProgress = True Then
                MainForm.ResumeGameLabel.Visible = True
                MainForm.NewGameLabel.Top = MMLLastPos(1)
                MainForm.SaveGameLabel.Top = MMLLastPos(2)
                MainForm.LoadGameLabel.Top = MMLLastPos(3)
                MainForm.OptionsLabel.Top = MMLLastPos(4)
                MainForm.ExitLabel.Top = MMLLastPos(5)
            Else
                MainForm.ResumeGameLabel.Visible = False
                MainForm.NewGameLabel.Top = MMLOrigPos(1)
                MainForm.SaveGameLabel.Top = MMLOrigPos(2)
                MainForm.LoadGameLabel.Top = MMLOrigPos(3)
                MainForm.OptionsLabel.Top = MMLOrigPos(4)
                MainForm.ExitLabel.Top = MMLOrigPos(5)
            End If
        Case "Options"
            MainForm.Width = MainForm.OptionFrame.Width + 100
            MainForm.Height = MainForm.OptionFrame.Height + 450
            MainForm.GameFrame.Visible = False
            MainForm.MainMenuFrame.Visible = False
            MainForm.OptionFrame.Visible = True
            MainForm.SaveGameFrame.Visible = False
            MainForm.LoadGameFrame.Visible = False
            Select Case SelectedSlot
                Case 1
                    MainForm.ASSlotOption(0).Value = 1
                Case 2
                    MainForm.ASSlotOption(1).Value = 1
                Case 3
                    MainForm.ASSlotOption(2).Value = 1
            End Select
            Select Case AutoSave
                Case 0
                    MainForm.AutoSaveOption(2).Value = 1
                Case 1
                    MainForm.AutoSaveOption(0).Value = 1
                Case 2
                    MainForm.AutoSaveOption(1).Value = 1
            End Select
        Case "GameFrame"
            MainForm.Width = MainForm.GameFrame.Width + 100
            MainForm.Height = MainForm.GameFrame.Height + 450
            MainForm.GameFrame.Visible = True
            MainForm.MainMenuFrame.Visible = False
            MainForm.OptionFrame.Visible = False
            MainForm.SaveGameFrame.Visible = False
            MainForm.LoadGameFrame.Visible = False
            GameInProgress = True
        Case "SaveGame"
            MainForm.Width = MainForm.SaveGameFrame.Width + 100
            MainForm.Height = MainForm.SaveGameFrame.Height + 450
            MainForm.GameFrame.Visible = False
            MainForm.MainMenuFrame.Visible = False
            MainForm.OptionFrame.Visible = False
            MainForm.SaveGameFrame.Visible = True
            MainForm.LoadGameFrame.Visible = False
        Case "LoadGame"
            MainForm.Width = MainForm.LoadGameFrame.Width + 100
            MainForm.Height = MainForm.LoadGameFrame.Height + 450
            MainForm.GameFrame.Visible = False
            MainForm.MainMenuFrame.Visible = False
            MainForm.OptionFrame.Visible = False
            MainForm.SaveGameFrame.Visible = False
            MainForm.LoadGameFrame.Visible = True
    End Select
    MainForm.NewGameLabel.ForeColor = &H8000000E
    MainForm.LoadGameLabel.ForeColor = &H8000000E
    MainForm.SaveGameLabel.ForeColor = &H8000000E
    MainForm.OptionsLabel.ForeColor = &H8000000E
    MainForm.ReturnToMainMenu(0).ForeColor = &H8000000E
    MainForm.ReturnToMainMenu(1).ForeColor = &H8000000E
    MainForm.ReturnToMainMenu(2).ForeColor = &H8000000E
    For i = 1 To MainForm.SlotLabel.Count
        MainForm.SlotLabel(i).ForeColor = &H8000000E
    Next i
End Sub

Sub UpdateAll()
    Call UpdateBerry(0, False)
    Call UpdateHealth(0, False)
    Call UpdateFurs(0, False)
    Call UpdateLeaf(0, False)
    Call UpdateLog(0, False)
    Call UpdateStick(0, False)
    Call UpdateStone(0, False)
    Call UpdateWater(0, False)
    Call UpdateInventory(0, False)
    Call UpdateHunger(0, False)
    Call UpdateEnergy(0, False)
    Call UpdateThirst(0, False)
    Call UpdateMeat(0, False)
    Call UpdateRawMeat(0, False)
    DayNum = DayNum - 1
    Call UpdateDay(False)
    Call UpdateEnergyPerNight
    Call UpdateThingsCooked(0, False)
    Call UpdateEnergySpent(0, False)
    Call UpdateAnimalsHunted(0, False)
    Call UpdateAllCollected(0, False)
    Call UpdateThingsCrafted(0, False)
    Call UpdateHealthGain(0, False)
    Call UpdateAllLevels(0, False)
    Call UpdateActionsCompleted(0, False)
    Call UpdateTraps(0, False)
    Call UpdateCWolf(False)
    Call UpdateCWolfXP(0, False)
    Call UpdateCWolfHealth(0, False)
    Call UpdateCWolfHunger(0, False)
    Call UpdateCWolfBond(0, False)
    Call UpdateCWolfLvl(0, False)
    Call UpdateWolfBond(0, False)
    Call UpdateArrows(0, False)
    Call UpdateVines(0, False)
    Call UpdateFeathers(0, False)
    Call UpdateRabbits(0, False)
    Call UpdateRabbitHunger(0, False)
    Call UpdateBirdTraps(0, False)
    Call UpdateLiveTraps(0, False)
    Call UpdateBow(False)
    Call UpdateRabbitCage(False)
    
End Sub

Sub UpdateCornPlanted(Amount As Integer, Reset As Boolean)
    If Reset = False Then
        CornPlanted = CornPlanted + Amount
        UpdateTotalCropsPlanted
    Else
        CornPlanted = 0
    End If
    MainForm.CornPlantedLabel.Caption = "Corn planted: " & CornPlanted
End Sub
