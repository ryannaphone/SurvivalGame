VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "CreativeGameName"
   ClientHeight    =   5100
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame TSFrame 
      Caption         =   "Tab Selection"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton CookButton 
         Caption         =   "Cooking"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CommandButton HuntButton 
         Caption         =   "Hunting"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CommandButton BuildButton 
         Caption         =   "Building"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton CraftButton 
         Caption         =   "Crafting"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton ConsumeButton 
         Caption         =   "Consume Items"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton GatherButton 
         Caption         =   "Gathering"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame LevelsFrame 
      Caption         =   "Level: 1"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5160
      TabIndex        =   50
      Top             =   1440
      Width           =   1815
      Begin VB.Label GathLvlLabel 
         Caption         =   "Gathering: 1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label CraftLvlLabel 
         Caption         =   "Crafting: 1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label HealLvlLabel 
         Caption         =   "Health: 1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label StrenLvlLabel 
         Caption         =   "Strength: 1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label HuntLvlLabel 
         Caption         =   "Hunting: 1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label CookLvlLabel 
         Caption         =   "Cooking: 1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label StamLvlLabel 
         Caption         =   "Stamina: 1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton BackButton 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Frame BuildingFrame 
      Caption         =   "Buildings"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3360
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
      Begin VB.Label SSLabel 
         Caption         =   "No Stick Shelter"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label FireLabel 
         Caption         =   "No Fire"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame InventoryFrame 
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   7200
      TabIndex        =   8
      Top             =   0
      Width           =   1815
      Begin VB.Label AxeLabel 
         Caption         =   "No Axe"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label MeatLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label RabbitLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label StoneLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label SpearLabel 
         Caption         =   "No Spear"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label WaterLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label LeafLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label LogLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label StickLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label BerryLabel 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton EndDayButton 
      Caption         =   "End Day"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      TabIndex        =   6
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Frame Stats 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.Label ThirstLabel 
         Caption         =   "Health: 100/100"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label HungerLabel 
         Caption         =   "Health: 100/100"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label HealthLabel 
         Caption         =   "Health: 100/100"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label EnergyLabel 
         Caption         =   "Health: 100/100"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame CraftFrame 
      Caption         =   "Crafting"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   31
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton AxeCraft 
         Caption         =   "Axe"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   1355
      End
      Begin VB.CommandButton StickCraft 
         Caption         =   "Sticks"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   57
         Top             =   840
         Width           =   1355
      End
      Begin VB.CommandButton SpearCraft 
         Caption         =   "Spear"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   1355
      End
      Begin VB.CommandButton FireCraft 
         Caption         =   "Fire"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   36
         Top             =   360
         Width           =   1355
      End
   End
   Begin VB.Frame BuildFrame 
      Caption         =   "Building"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   32
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton SSBuild 
         Caption         =   "Stick Shelter"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame CIFrame 
      Caption         =   "Consume Items"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton MeatConsume 
         Caption         =   "Meat"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   47
         Top             =   360
         Width           =   1355
      End
      Begin VB.CommandButton BerryConsume 
         Caption         =   "Berry"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   1355
      End
      Begin VB.CommandButton WaterConsume 
         Caption         =   "Water"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1355
      End
   End
   Begin VB.Frame GRFrame 
      Caption         =   "Gather Resources"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton GatherStones 
         Caption         =   "Stones"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   1355
      End
      Begin VB.CommandButton GatherLeaves 
         Caption         =   "Leaves"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1355
      End
      Begin VB.CommandButton GatherLogs 
         Caption         =   "Logs"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1355
      End
      Begin VB.CommandButton GatherWater 
         Caption         =   "Water"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   1355
      End
      Begin VB.CommandButton GatherSticks 
         Caption         =   "Sticks"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton GatherBerry 
         Caption         =   "Berries"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1355
      End
   End
   Begin VB.Frame HuntFrame 
      Caption         =   "Hunting"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   34
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton RabbitHunt 
         Caption         =   "Rabbit"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1355
      End
   End
   Begin VB.Frame CookFrame 
      Caption         =   "Cooking"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   43
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton RabbitCook 
         Caption         =   "Rabbit"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1355
      End
   End
   Begin VB.Line Line2 
      X1              =   7080
      X2              =   7080
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Label NewsLabel 
      Caption         =   "News: Heyo! Imagine some news here, yea?"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7080
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Label InfoLabel 
      Caption         =   "Info: Enjoy the game!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label DayLabel 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu MenAbt 
      Caption         =   "About"
   End
   Begin VB.Menu MenHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AxeCraft_Click()
    CraftAxe
End Sub

Private Sub AxeCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Craft an axe using 1 stick and 2 stones. The axe allows you to get more logs."
End Sub

Private Sub Form_Load()
    RestartGame
End Sub

Private Sub BackButton_Click() 'Takes the user back to the tab selection frame by making all others invisible
    TabSelect = True
    BackButton.Visible = False
    GRFrame.Visible = False
    TSFrame.Visible = True
    CIFrame.Visible = False
    CraftFrame.Visible = False
    BuildFrame.Visible = False
    HuntFrame.Visible = False
    CookFrame.Visible = False
End Sub

Private Sub GatherButton_Click() 'Brings up the gathering tab
    TabSelect = False
    BackButton.Visible = True
    GRFrame.Visible = True
    TSFrame.Visible = False
    CIFrame.Visible = False
    CraftFrame.Visible = False
    BuildFrame.Visible = False
    HuntFrame.Visible = False
End Sub

Private Sub BuildButton_Click() 'Brings up the building tab
    TabSelect = False
    BackButton.Visible = True
    GRFrame.Visible = False
    TSFrame.Visible = False
    CIFrame.Visible = False
    CraftFrame.Visible = False
    BuildFrame.Visible = True
    HuntFrame.Visible = False
End Sub

Private Sub CraftButton_Click() 'Brings up the crafting tab
    TabSelect = False
    BackButton.Visible = True
    GRFrame.Visible = False
    TSFrame.Visible = False
    CIFrame.Visible = False
    CraftFrame.Visible = True
    BuildFrame.Visible = False
    HuntFrame.Visible = False
End Sub

Private Sub ConsumeButton_Click() 'Brings up the consume items tab
    TabSelect = False
    BackButton.Visible = True
    GRFrame.Visible = False
    TSFrame.Visible = False
    CIFrame.Visible = True
    CraftFrame.Visible = False
    BuildFrame.Visible = False
    HuntFrame.Visible = False
End Sub

Private Sub HuntButton_Click() 'Brings up the hunting tab
    TabSelect = False
    BackButton.Visible = True
    GRFrame.Visible = False
    TSFrame.Visible = False
    CIFrame.Visible = False
    CraftFrame.Visible = False
    BuildFrame.Visible = False
    HuntFrame.Visible = True
End Sub

Private Sub CookButton_Click() 'Brings up the cooking tab
    If Fire = True Then
        TabSelect = False
        BackButton.Visible = True
        GRFrame.Visible = False
        TSFrame.Visible = False
        CIFrame.Visible = False
        CraftFrame.Visible = False
        BuildFrame.Visible = False
        HuntFrame.Visible = False
        CookFrame.Visible = True
    Else
        NewsLabel.Caption = "News: You need a fire to cook."
    End If
End Sub

Private Sub EndDayButton_Click() 'Ends the day
    If Energy > 0 Then
        If MsgBox("You have unspent energy, are you sure you want to end the day?", vbYesNo, "Are you sure?") = vbYes Then
            EndDay
        Else
            Exit Sub
        End If
    Else
        EndDay
    End If
End Sub

Private Sub GatherBerry_Click()
    If Energy >= 5 Then
        BerryGather
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherLeaves_Click()
    If Energy >= 2 Then
        LeafGather
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherLogs_Click()
    If Energy >= 8 Then
        LogGather
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherSticks_Click()
    If Energy >= 5 Then
        StickGather
        'Call UpdateEnergy(-5, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherWater_Click()
    If Energy >= 5 Then
        WaterGather
        'Call UpdateEnergy(-5, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherStones_Click()
    If Energy >= 5 Then
        StoneGather
        'Call UpdateEnergy(-5, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub BerryConsume_Click()
    If Berries > 0 Then
        ConsumeBerry
        NewsLabel.Caption = "News: You ate a berry."
    Else
        NewsLabel.Caption = "News: You don't have a berry to eat."
    End If
End Sub

Private Sub MenAbt_Click()
    MsgBox "SurvivalGameThing? V0.1 Made by Ryan Watson on 1/17/2016."
End Sub

Private Sub MenHelp_Click()
    MsgBox "The goal of the game is to survive. You do so by keeping your health, hunger, and thirst above 0. If your health reaches 0 you lose and if your hunger or thirst reach 0 you will begin to lose health. Spend energy to obtain resources and then proceed to the next day. Good luck and enjoy!"
End Sub

Private Sub MeatConsume_Click()
    If Meat > 0 Then
        ConsumeMeat
        NewsLabel.Caption = "News: You ate some meat."
    Else
        NewsLabel.Caption = "News: You don't have any meat."
    End If
End Sub

Private Sub RabbitCook_Click()
    If Rabbits > 0 Then
        CookRabbit
        Call UpdateThingsCooked(1, False)
    Else
        NewsLabel.Caption = "News: You don't have a rabbit to cook."
    End If
End Sub

Private Sub RabbitHunt_Click()
    If Energy >= 6 Then
        HuntRabbit
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub SpearCraft_Click()
    If Spear = False Then
        CraftSpear
    End If
End Sub

Private Sub FireCraft_Click()
    If Fire = False Then
        CraftFire
    End If
End Sub

Private Sub SSBuild_Click()
    If StickShelt = False Then
        BuildSShelter
    End If
End Sub

Private Sub StickCraft_Click()
    CraftSticks
End Sub

Private Sub StickCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Turn a log into sticks! Requires 1 log."
End Sub

Private Sub WaterConsume_Click()
    ConsumeWater
End Sub

Private Sub HealthLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: If your health reaches 0, you die. Your health increases slowly over time or you can take medicine to increase it."
End Sub

Private Sub HungerLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: If your hunger reaches 0, you will begin to lose health. You can increase your hunger by eating food."
End Sub

Private Sub InventoryFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: This displays the amount of each resource that you have."
End Sub

Private Sub BuildingFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: This displays the amount of each building that you have."
End Sub

Private Sub GRFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: You can gather various resources by clicking the buttons in this tab."
End Sub

Private Sub ThirstLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: If your thirst reaches 0, you will begin to lose health. You can increase your thirst by drinking water."
End Sub

Private Sub EndDay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to end the day and go to sleep. You will regain energy while you sleep."
End Sub

Private Sub EnergyLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Energy is used to complete actions. You regain energy by eating, drinking, or sleeping(ending the day.)"
End Sub

Private Sub GatherBerry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to gather berries. It costs 5 energy."
End Sub

Private Sub GatherLeaves_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to gather leaves. It costs 2 energy."
End Sub

Private Sub GatherLogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to gather logs. It costs 8 energy."
End Sub

Private Sub GatherWater_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to gather water. It costs 5 energy."
End Sub

Private Sub GatherSticks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to gather sticks. It costs 5 energy."
End Sub

Private Sub BerryLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: You can eat berries to regain 1 energy, 1 hunger, 1 thirst, and 1 health. Each berry takes up 1 inventory space."
End Sub

Private Sub WaterLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: You can drink water to regain 1 energy, 2 thirst, and 1 health. Each cup takes up 1 inventory space."
End Sub

Private Sub LogLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: You can use logs for building and crafting. Each log takes up 5 inventory space."
End Sub

Private Sub LeafLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: You can use leaves for building and crafting. Leaves take up no inventory space."
End Sub

Private Sub StickLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: You can use sticks for building and crafting. Each stick takes up 1 inventory space."
End Sub

Private Sub BackButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to return to tab selection."
End Sub

Private Sub FireCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to make a fire. You need 3 log, 10 sticks, and 25 leaves."
End Sub

Private Sub SpearCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to make a spear. You need 1 sticks, and 1 stone."
End Sub

Private Sub FireLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: A fire allows you to cook food. Go to the crafting tab for more information."
End Sub

Private Sub SSBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click to build a stick shelter. It requires 25 sticks and 50 leaves."
End Sub

Private Sub SSLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: The stick shelter increases the amount of rest(energy) that you get each night."
End Sub

Private Sub SpearLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: The spear increases your chances of catching something while hunting."
End Sub

Private Sub MeatLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Cooked meat can be eaten for a higher amount of energy and hunger than berries."
End Sub

Private Sub RabbitLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Uncooked rabbit meat cannot be eaten, but it can be cooked."
End Sub

Private Sub AxeLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: The axe allows you to chop down trees and craft more stuff. Requires 1 stick and 2 stones."
End Sub

Private Sub StoneLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Stones are used for crafting and building."
End Sub

Sub NotEnoughEnergy()
    NewsLabel.Caption = "News: You don't have enough energy to complete this action."
End Sub
