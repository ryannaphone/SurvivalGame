Attribute VB_Name = "SGVariables"
Global Health As Integer 'Amount of health currently held by the player
Global Hunger As Integer 'Amount of hunger currently held by the player
Global Thirst As Integer 'Amount of thirst currently held by the player
Global Energy As Integer 'Amount of energy currently held by the player
Global MaxHealth As Integer 'Maximum amount of health the player can have
Global MaxEnergy As Integer 'Maximum amount of energy the player can have at one time
Global MaxHunger As Integer 'Maximum amount of hunger the player can have at one time (more is better)
Global MaxThirst As Integer 'Maximum amount of thirst the player can have at one time (more is better)
Global DayNum As Integer 'Day number (starts at 1)
Global Berries As Integer 'Amount of berries held in inventory
Global Water As Integer 'Amount of water held in inventory (in cups)
Global Sticks As Integer 'Amount of sticks held in inventory
Global Logs As Integer 'Amount of logs held in inventory
Global Leaves As Integer 'Amount of leaves held in inventory
Global Stones As Integer 'Amount of stones held in inventory
Global Inventory As Integer 'Currently used amount of inventory space
Global MaxInventory As Integer 'Maximum amount of inventory space
Global TabSelect As Boolean 'Unused, but if true then the tab selection frame is on top and visible
Global Spear As Boolean 'If false player doesn't have a spear
Global Fire As Boolean 'If false the player doesn't have a fire
Global Meat As Integer 'Amount of cooked meat in inventory
Global EnergyPerNight As Integer 'Amount of energy regained per night
Global StickShelt As Boolean 'If false, then the player doesn't have a stick shelter
Global ThingsCooked As Integer 'Amount of things cooked successfully, influences cooking level
Global HuntFailed As Integer 'If =3 then the hunt failed (used to tell whether or not hunt failed)
Global AnimalsHunted As Integer 'Amount of animals successfully hunted
Global HuntingLvl As Integer 'Level 1 to 10 for hunting, influences odds of success
Global CookingLvl As Integer 'Level 1 to 10 for cooking, influences odds of success
Global StrengthLvl As Integer 'Level 1 to 10 for strength, reduces amount of energy required for certain actions
Global GatheringLvl As Integer 'Level 1 to 10 for gathering, influences odds of success and maybe amounts found
Global StaminaLvl As Integer 'Level 1 to 10 for stamina, influences max energy
Global PlayerLvl As Integer 'Level 1 to 10 for general player stats
Global HealthLvl As Integer 'Level 1 to 10 for health/vitality, influences max health and damage taken
Global StonesCollected As Integer 'Amount of stones collected throughout the entire game, influences gathering
Global SticksCollected As Integer 'Amount of sticks collected throughout the entire game, influences gathering
Global WaterCollected As Integer 'Amount of water collected throughout the entrire game, influences gathering
Global BerriesCollected As Integer 'Amount of berries collected throughout the entire game, influences gathering
Global LogsCollected As Integer 'Amount of logs collected throughout the entire game, influences gathering
Global LeavesCollected As Integer 'Amount of leaves collected throughout the entire game, influences gathering
Global AllCollected As Integer 'Amount of everything collected, influences gathering level
Global Axe As Boolean 'If true then the user has crafted the axe
Global CraftingLevel As Integer 'Level 1 to 10 that determines amount of stuff required for crafting stuff
Global LogsPerGather As Integer 'The amount of possible logs one can get per gather
Global RangeS As Integer 'Starting point for ranges
Global RangeE As Integer 'Ending point for ranges
Global HuntThresh As Integer 'The point at which the players hunting level increases
Global PlayerThresh As Integer 'The point at which overall level increases
Global HealthThresh As Integer 'the point at which the players health level increases
Global StaminaThresh As Integer 'The point at which stamina level increases
Global GathThresh As Integer 'The point at which gathering level increases
Global StrengthThresh As Integer 'The point at which strength level increases
Global CookThresh As Integer 'The point at which cooking level increases
Global CraftThresh As Integer 'The point at which crafting level increases
Global ThingsCrafted As Integer 'The amount of things crafted
Global EnergySpent As Integer 'The amount of energy used (negative)
Global HealthGain As Integer 'Amount of health gained
Global Temp1 As Integer 'random variable
Global Cabin As Boolean 'if true then user has built a cabin and energypernight = 20
Global Bed As Boolean 'if true then the user has built a bed and energypernight = 25
Global Furs As Integer 'used to craft a bed and gained by skinning rabbits
Global Knife As Boolean 'if true then the user has a knife and can use it for crafting and skinning
Global Temperature As Integer
Global Warmth As Integer
Global Time As Integer
Global RawMeat As Integer 'Amount of uncooked meat in the player's inventory
Global ConsoleOpen As Boolean 'Returns whether console is open or not
Global AllLevels As Integer
Global ActionsCompleted As Integer
Global SnipeEnabled As Boolean
Global Traps As Integer
Global FireStartDay As Integer
Global SetTraps As Integer
Global CWolf As Boolean 'Returns whether or not the user has a wolf companion
Global CWolfLvl As Integer 'Returns the level of the wolf companion
Global CWolfHealth As Integer 'Returns the health of the wolf companion
Global CWolfHunger As Integer 'Returns the hunger of the wolf companion
Global CWolfBond As Integer 'Returns the bond amount of the wolf companion
Global CWolfBondLvl As Integer 'Returns the bond level of the wolf companion
Global CWolfMaxHealth As Integer
Global CWolfMaxHunger As Integer
Global CWolfXP As Integer 'Amount of experience, influences CWolfLvl
Global CWolfThresh As Integer
Global CWolfBondThresh As Integer
Global CWolfDailyAction As Boolean 'Returns whether or not the wolf has done its daily action
Global LastCWolfBondThresh As Integer 'The last threshold for increasing wolf bond level
Global LastCommand As String 'Returns the last cheat command entered by the user
Global WolfPrompt1 As Boolean
Global WolfBond As Integer 'Bond amount before the wolf becomes the user's companion
Global WolfPrompt2 As Boolean 'All six wolfprompt booleans are used for determining which situation prompt will appear next
Global WolfPrompt3 As Boolean
Global CWolfName As String 'The changable name of the wolf companion
Global Arrows As Integer 'Amount of arrows that the user has in inventory
Global Vines As Integer 'Amount of vines that the user has in inventory
Global Bow As Boolean 'Returns whether or not the user has crafted a bow
Global LTraps As Integer 'Amount of live traps that the user has
Global LTrapsSet As Integer
Global BTraps As Integer 'Amount of bird traps that the user has
Global BTrapsSet As Integer
Global RCage As Boolean 'Returns whether or not the user has built a rabbit cage
Global Rabbits As Integer 'Amount of live rabbits held in cage
Global Feathers As Integer 'Amount of feathers in the user's inventory
Global MaxRabbits As Integer 'Maximum amount of rabbits the user can have in their cage
Global RabbitHunger As Integer 'Current rabbit hunger
Global MaxRabbitHunger As Integer 'Maximum amount of hunger the rabbits can have
Global Stats() As String 'Stats to be saved and recalled
Global StatNum As Integer
Global AutoSave As Integer 'If 0 then never autosave, if 1 then autosave on exit, if 2 then everyday
Global SelectedSlot As Integer
Global LoadedGame As Boolean
Global GameInProgress As Boolean 'Returns whether or not a game is in progress
Global MMLLastPos(1 To 5) As Integer
Global MMLOrigPos(1 To 5) As Integer
Global Corn As Integer
Global CornSeeds As Integer
Global CornPlanted As Integer
Global TotalCropsPlanted As Integer
Global BerryBushesPlanted As Integer
Global ClickValue As Integer
