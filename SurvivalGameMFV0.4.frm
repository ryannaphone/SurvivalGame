VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CreativeGameName"
   ClientHeight    =   6675
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame GameFrame 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Frame BuildMenuFrame 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4815
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   6375
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
            Height          =   2295
            Left            =   3240
            TabIndex        =   48
            Top             =   1800
            Width           =   3135
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
               TabIndex        =   55
               Top             =   360
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
               TabIndex        =   54
               Top             =   840
               Width           =   1335
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
               TabIndex        =   53
               Top             =   1320
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
               TabIndex        =   52
               Top             =   840
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
               TabIndex        =   51
               Top             =   360
               Width           =   1355
            End
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
               TabIndex        =   50
               Top             =   1320
               Width           =   1355
            End
            Begin VB.CommandButton VineGather 
               Caption         =   "Vines"
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
               TabIndex        =   49
               Top             =   1800
               Width           =   1335
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
            Height          =   2295
            Left            =   0
            TabIndex        =   39
            Top             =   1320
            Width           =   3135
            Begin VB.CommandButton SSBuild 
               Caption         =   "Stick Shelter"
               BeginProperty Font 
                  Name            =   "Rockwell"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   47
               Top             =   1800
               Width           =   1335
            End
            Begin VB.CommandButton CabinBuild 
               Caption         =   "Log Cabin"
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
               TabIndex        =   46
               Top             =   1320
               Width           =   1335
            End
            Begin VB.CommandButton BedBuild 
               Caption         =   "Bed"
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
               TabIndex        =   45
               Top             =   360
               Width           =   1335
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
               Left            =   120
               TabIndex        =   44
               Top             =   840
               Width           =   1335
            End
            Begin VB.CommandButton TrapBuild 
               Caption         =   "Traps"
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
               TabIndex        =   43
               Top             =   1800
               Width           =   1335
            End
            Begin VB.CommandButton RCageBuild 
               Caption         =   "Rabbit Cage"
               BeginProperty Font 
                  Name            =   "Rockwell"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   42
               Top             =   1320
               Width           =   1355
            End
            Begin VB.CommandButton BTrapBuild 
               Caption         =   "Bird Trap"
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
               TabIndex        =   41
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton LTrapBuild 
               Caption         =   "Live Traps"
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
               TabIndex        =   40
               Top             =   840
               Width           =   1335
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
            Height          =   1815
            Left            =   3240
            TabIndex        =   32
            Top             =   0
            Width           =   3135
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
               TabIndex        =   38
               Top             =   1320
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
               TabIndex        =   37
               Top             =   1320
               Width           =   1355
            End
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
               Left            =   1680
               TabIndex        =   36
               Top             =   360
               Width           =   1355
            End
            Begin VB.CommandButton KnifeCraft 
               Caption         =   "Knife"
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
               TabIndex        =   35
               Top             =   840
               Width           =   1355
            End
            Begin VB.CommandButton BowCraft 
               Caption         =   "Bow"
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
               TabIndex        =   34
               Top             =   840
               Width           =   1355
            End
            Begin VB.CommandButton ArrowCraft 
               Caption         =   "Arrows"
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
               TabIndex        =   33
               Top             =   360
               Width           =   1355
            End
         End
         Begin VB.Frame MaintenanceFrame 
            Caption         =   "Maintenance"
            BeginProperty Font 
               Name            =   "Rockwell"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   3135
            Begin VB.CommandButton FireMaintain 
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
               TabIndex        =   31
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton TrapMaintain 
               Caption         =   "Traps"
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
               TabIndex        =   30
               Top             =   840
               Width           =   1335
            End
            Begin VB.CommandButton BTrapMaintain 
               Caption         =   "Bird Traps"
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
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton LTrapMaintain 
               Caption         =   "Live Traps"
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
               TabIndex        =   28
               Top             =   840
               Width           =   1335
            End
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Animal Menu"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton StatMenuButton 
         Caption         =   "Stat Menu"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton FoodMenuButton 
         Caption         =   "Food Menu"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton BuildMenuButton 
         Caption         =   "Build Menu"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1455
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
         Height          =   5535
         Left            =   6720
         TabIndex        =   5
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton ButtonValue 
            Caption         =   "*-1"
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
            Index           =   3
            Left            =   960
            TabIndex        =   9
            Top             =   5160
            Width           =   615
         End
         Begin VB.CommandButton ButtonValue 
            Caption         =   "*100"
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
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   5160
            Width           =   735
         End
         Begin VB.CommandButton ButtonValue 
            Caption         =   "*10"
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
            Index           =   1
            Left            =   600
            TabIndex        =   7
            Top             =   4800
            Width           =   615
         End
         Begin VB.CommandButton ButtonValue 
            Caption         =   "*1"
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
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   4800
            Width           =   375
         End
         Begin VB.Label FeatherLabel 
            Caption         =   "0 Feathers"
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
            TabIndex        =   25
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label ArrowLabel 
            Caption         =   "0 Arrows"
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
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label BowLabel 
            Caption         =   "No Bow"
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
            TabIndex        =   23
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label VineLabel 
            Caption         =   "0 Vines"
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
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label FurLabel 
            Caption         =   "0 Furs"
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
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label KnifeLabel 
            Caption         =   "No Knife"
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
            Top             =   3600
            Width           =   1455
         End
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
            TabIndex        =   19
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label MeatLabel 
            Caption         =   "0 Meat"
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
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label RawMeatLabel 
            Caption         =   "0 Raw Meat"
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
            TabIndex        =   17
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label StoneLabel 
            Caption         =   "0 Stones"
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
            TabIndex        =   16
            Top             =   2640
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
            TabIndex        =   15
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label WaterLabel 
            Caption         =   "0 Cups of Water"
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
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label LeafLabel 
            Caption         =   "0 Leaves"
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
            TabIndex        =   13
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label LogLabel 
            Caption         =   "0 Logs"
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
            TabIndex        =   12
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label StickLabel 
            Caption         =   "0 Sticks"
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
            TabIndex        =   11
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label BerryLabel 
            Caption         =   "0 Berries"
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
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Line Line2 
         X1              =   6600
         X2              =   0
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   6600
         X2              =   6600
         Y1              =   0
         Y2              =   5640
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AnimalFeed_Click()
    If MainForm.AnimalActFrame.Caption = "Rabbit - Actions" Then
        FeedRabbits
    Else
        MainForm.NewsLabel.Caption = "News: Select an animal first."
    End If
End Sub

Private Sub AnimalFeed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Feed the selected animal."
End Sub

Private Sub AnimalFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click on the animals in this form to select one to manage."
End Sub

Private Sub AnimalKill_Click()
    If MainForm.AnimalActFrame.Caption = "Rabbit - Actions" Then
        KillRabbit
    Else
        MainForm.NewsLabel.Caption = "News: Select an animal first."
    End If
End Sub

Private Sub AnimalKill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Kill one of the selected animal."
End Sub

Private Sub ArrowCraft_Click()
    CraftArrows
End Sub

Private Sub ArrowCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to craft 5 arrows. It requires 5 sticks, 5 stones, and 5 feathers."
End Sub

Private Sub ASSlotOption_Click(Index As Integer)
    Select Case Index
        Case 0
            ASSlotOption(1).Value = 0
            ASSlotOption(2).Value = 0
            SelectedSlot = 1
        Case 1
            ASSlotOption(0).Value = 0
            ASSlotOption(2).Value = 0
            SelectedSlot = 2
        Case 2
            ASSlotOption(1).Value = 0
            ASSlotOption(0).Value = 0
            SelectedSlot = 3
    End Select
End Sub

Private Sub AutoSaveOption_Click(Index As Integer)
    Select Case Index
        Case 0
            AutoSaveOption(1).Value = 0
            AutoSaveOption(2).Value = 0
            AutoSave = 1
        Case 1
            AutoSaveOption(0).Value = 0
            AutoSaveOption(2).Value = 0
            AutoSave = 2
        Case 2
            AutoSaveOption(1).Value = 0
            AutoSaveOption(0).Value = 0
            AutoSave = 0
    End Select
End Sub

Private Sub AxeCraft_Click()
    CraftAxe
    Call UpdateActionsCompleted(1, False)
End Sub

Private Sub AxeCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Craft an axe using 1 stick and 2 stones. The axe allows you to get more logs."
End Sub

Private Sub BackButton2_Click()
    BackButton2.Visible = False
    WolfCommandFrame.Visible = False
    CompActFrame.Visible = True
    WolfInteractFrame.Visible = False
End Sub

Private Sub BearHunt_Click()
    If Energy >= 10 Then
        HuntBear
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub BearHunt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to hunt a bear. Costs 10 energy."
End Sub

Private Sub BedBuild_Click()
    BuildBed
    Call UpdateActionsCompleted(1, False)
End Sub

Private Sub BedBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Requires cabin. Use 6 logs, 20 sticks, and 25 furs to make a bed. Increases rest(energy) per night."
End Sub

Private Sub BowCraft_Click()
    CraftBow
End Sub

Private Sub BowCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to craft a bow. Requires 3 sticks and 2 vines."
End Sub

Private Sub BTrapBuild_Click()
    CraftBirdTrap
End Sub

Private Sub BTrapBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to build a bird trap. Requires 10 sticks, 3 berries, and 5 vines."
End Sub

Private Sub BTrapLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     InfoLabel.Caption = "Info: Bird traps have a chance to catch birds if they are set. You have " & BTrapsSet & " bird traps set."
End Sub

Private Sub BTrapMaintain_Click()
    MaintainBirdTraps
End Sub

Private Sub BTrapMaintain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to reset one of your bird traps. Costs 3 berries."
End Sub

Private Sub ButtonValue_Click(Index As Integer)
    Select Case Index
        Case 0
            ClickValue = 1
        Case 1
            ClickValue = 10
        Case 2
            ClickValue = 100
        Case 3
            ClickValue = -1
    End Select
End Sub

Private Sub CabinBuild_Click()
    BuildCabin
    Call UpdateActionsCompleted(1, False)
End Sub

Private Sub CabinBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: USe 50 logs and 50 sticks to build a log cabin. Increases rest(energy) per night."
End Sub

Private Sub CabinLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: The cabin increases your rest(energy) per night."
End Sub

Private Sub BedLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: The bed increases your rest(energy) per night. You must build a cabin first."
End Sub

Private Sub CommandsButton_Click()
    If CompActFrame.Caption = "Wolf - Actions" Then
        WolfCommandFrame.Visible = True
        CompActFrame.Visible = False
        WolfInteractFrame.Visible = False
        BackButton2.Visible = True
    End If
    If CompActFrame.Caption = "Select a Companion" Then
        NewsLabel.Caption = "News: You must select a companion first."
    Else
        
    End If
End Sub

Private Sub CompanionFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click on the names in this frame to select a companion."
End Sub

Private Sub ExitLabel_Click()
    If AutoSave = 1 Then SaveGame (SelectedSlot)
    End
End Sub

Private Sub ExitLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewGameLabel.ForeColor = &H8000000E
    LoadGameLabel.ForeColor = &H8000000E
    SaveGameLabel.ForeColor = &H8000000E
    OptionsLabel.ForeColor = &H8000000E
    ExitLabel.ForeColor = &H8000000A
    ResumeGameLabel.ForeColor = &H8000000E
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'MsgBox KeyCode
    Select Case KeyCode
        Case 27
            Call UpdateOpenFrame("MainMenu")
        Case 38
            Console.Text = LastCommand
    End Select
End Sub

Private Sub InteractButton_Click()
    If CompActFrame.Caption = "Wolf - Actions" Then
        WolfInteractFrame.Visible = True
    End If
    If CompActFrame.Caption = "Select a Companion" Then
        NewsLabel.Caption = "News: You must select a companion first."
    Else
    CompActFrame.Visible = False
    WolfCommandFrame.Visible = False
    BackButton2.Visible = True
    End If
End Sub

Private Sub CookLvlLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Your cooking level increases the odds of successfully cooking food."
End Sub

Private Sub CraftLvlLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Your crafting level is pretty cool and not at all arbitrary, right?"
End Sub

Private Sub HuntLvlLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Your hunting level increases your chances of successfully hunting animals."
End Sub

Private Sub HealLvlLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Your health level increases your max health."
End Sub

Private Sub GathLvlLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Your gathering level increases your chances of finding resources when gathering."
End Sub

Private Sub LoadGameLabel_Click()
    Call UpdateOpenFrame("LoadGame")
End Sub

Private Sub LoadGameLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewGameLabel.ForeColor = &H8000000E
    LoadGameLabel.ForeColor = &H8000000A
    SaveGameLabel.ForeColor = &H8000000E
    OptionsLabel.ForeColor = &H8000000E
    ExitLabel.ForeColor = &H8000000E
    ResumeGameLabel.ForeColor = &H8000000E
End Sub

Private Sub LTrapBuild_Click()
    CraftLiveTrap
End Sub

Private Sub LTrapBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to build a trap that catches live rabbits. Requires 10 sticks, 3 berries, and 5 vines."
End Sub

Private Sub LTrapLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     InfoLabel.Caption = "Info: Live traps have a chance to catch live rabbits if they are set. You have " & LTrapsSet & " live traps set."
End Sub

Private Sub LTrapMaintain_Click()
    MaintainLiveTraps
End Sub

Private Sub LTrapMaintain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click to reset one of your live rabbit traps. Costs 3 berries."
End Sub

Private Sub NewGameLabel_Click()
    Call UpdateOpenFrame("GameFrame")
    Call RestartGame
End Sub

Private Sub NewGameLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewGameLabel.ForeColor = &H8000000A
    LoadGameLabel.ForeColor = &H8000000E
    SaveGameLabel.ForeColor = &H8000000E
    OptionsLabel.ForeColor = &H8000000E
    ExitLabel.ForeColor = &H8000000E
    ResumeGameLabel.ForeColor = &H8000000E
End Sub

Private Sub OptionsLabel_Click()
    Call UpdateOpenFrame("Options")
End Sub

Private Sub OptionsLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewGameLabel.ForeColor = &H8000000E
    LoadGameLabel.ForeColor = &H8000000E
    SaveGameLabel.ForeColor = &H8000000E
    OptionsLabel.ForeColor = &H8000000A
    ExitLabel.ForeColor = &H8000000E
    ResumeGameLabel.ForeColor = &H8000000E
End Sub

Private Sub RabbitLabel_Click()
    AnimalActFrame.Caption = "Rabbit - Actions"
    MainForm.AnimalStats.Caption = "Rabbit - Stats"
    Call UpdateRabbitHunger(0, False)
End Sub

Private Sub RCageBuild_Click()
    BuildRabbitCage
End Sub

Private Sub RCageBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to build a rabbit cage. It requires 25 sticks and 10 stones."
End Sub

Private Sub RenameCompButton_Click()
    Load RenameForm
    RenameForm.Show
End Sub

Private Sub RenameCompButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to rename any companion."
End Sub

Private Sub ResumeGameLabel_Click()
    Call UpdateOpenFrame("GameFrame")
End Sub

Private Sub ResumeGameLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewGameLabel.ForeColor = &H8000000E
    LoadGameLabel.ForeColor = &H8000000E
    SaveGameLabel.ForeColor = &H8000000E
    OptionsLabel.ForeColor = &H8000000E
    ExitLabel.ForeColor = &H8000000E
    ResumeGameLabel.ForeColor = &H8000000A
End Sub

Private Sub ReturnToMainMenu_Click(Index As Integer)
    If Index = 0 Then
        If AutoSave <> 0 And SelectedSlot = 0 Then
            MsgBox "Please select a slot for autosaving."
            Exit Sub
        End If
    End If
    Call UpdateOpenFrame("MainMenu")
End Sub

Private Sub ReturnToMainMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReturnToMainMenu(Index).ForeColor = &H8000000A
End Sub

Private Sub SaveGameLabel_Click()
    Call UpdateOpenFrame("SaveGame")
End Sub

Private Sub SaveGameLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewGameLabel.ForeColor = &H8000000E
    LoadGameLabel.ForeColor = &H8000000E
    SaveGameLabel.ForeColor = &H8000000A
    OptionsLabel.ForeColor = &H8000000E
    ExitLabel.ForeColor = &H8000000E
    ResumeGameLabel.ForeColor = &H8000000E
End Sub

Private Sub SlotLabel_Click(Index As Integer)
    Select Case Index
        Case 1
            Call SaveGame(1)
        Case 2
            Call SaveGame(2)
        Case 3
            Call SaveGame(3)
        Case 4
            Call LoadGame(1)
        Case 5
            Call LoadGame(2)
        Case 6
            Call LoadGame(3)
    End Select
End Sub

Private Sub SlotLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 1
            SlotLabel(1).ForeColor = &H8000000A
            SlotLabel(2).ForeColor = &H8000000E
            SlotLabel(3).ForeColor = &H8000000E
        Case 2
            SlotLabel(1).ForeColor = &H8000000E
            SlotLabel(2).ForeColor = &H8000000A
            SlotLabel(3).ForeColor = &H8000000E
        Case 3
            SlotLabel(1).ForeColor = &H8000000E
            SlotLabel(2).ForeColor = &H8000000E
            SlotLabel(3).ForeColor = &H8000000A
        Case 4
            SlotLabel(4).ForeColor = &H8000000A
            SlotLabel(5).ForeColor = &H8000000E
            SlotLabel(6).ForeColor = &H8000000E
        Case 5
            SlotLabel(4).ForeColor = &H8000000E
            SlotLabel(5).ForeColor = &H8000000A
            SlotLabel(6).ForeColor = &H8000000E
        Case 6
            SlotLabel(4).ForeColor = &H8000000E
            SlotLabel(5).ForeColor = &H8000000E
            SlotLabel(6).ForeColor = &H8000000A
    End Select
End Sub

Private Sub StamLvlLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Your stamina level increases your max energy."
End Sub

Private Sub StrenLvlLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Your strength level increases your max inventory."
End Sub

Private Sub DeerHunt_Click()
    If Energy >= 8 Then
        HuntDeer
        Call UpdateActionsCompleted(1, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub DeerHunt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click to hunt a deer. They are harder to catch than rabbits and require more energy to hunt, but provide more meat."
End Sub

Private Sub FireMaintain_Click()
    MaintainFire
End Sub

Private Sub FireMaintain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Use two logs to make your fire last for 5 more days."
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'MsgBox KeyAscii
End Sub

Private Sub Form_Load()
    MMLLastPos(1) = MainForm.NewGameLabel.Top + MainForm.NewGameLabel.Height
    MMLLastPos(2) = MainForm.SaveGameLabel.Top + MainForm.SaveGameLabel.Height
    MMLLastPos(3) = MainForm.LoadGameLabel.Top + MainForm.LoadGameLabel.Height
    MMLLastPos(4) = MainForm.OptionsLabel.Top + MainForm.OptionsLabel.Height
    MMLLastPos(5) = MainForm.ExitLabel.Top + MainForm.ExitLabel.Height
    MMLOrigPos(1) = MainForm.NewGameLabel.Top
    MMLOrigPos(2) = MainForm.SaveGameLabel.Top
    MMLOrigPos(3) = MainForm.LoadGameLabel.Top
    MMLOrigPos(4) = MainForm.OptionsLabel.Top
    MMLOrigPos(5) = MainForm.ExitLabel.Top
    Call UpdateOpenFrame("MainMenu")
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

Private Sub FurLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Furs are gained by catching animals and can be used for crafting."
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

Private Sub MaintenanceButton_Click() 'Brings up the maintenance tab
    TabSelect = False
    BackButton.Visible = True
    GRFrame.Visible = False
    TSFrame.Visible = False
    CIFrame.Visible = False
    CraftFrame.Visible = False
    BuildFrame.Visible = False
    HuntFrame.Visible = False
    CookFrame.Visible = False
    MaintenanceFrame.Visible = True
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
        Call UpdateActionsCompleted(1, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherLeaves_Click()
    If Energy >= 2 Then
        LeafGather
        Call UpdateActionsCompleted(1, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherLogs_Click()
    If Energy >= 8 Then
        LogGather
        Call UpdateActionsCompleted(1, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherSticks_Click()
    If Energy >= 5 Then
        StickGather
        Call UpdateActionsCompleted(1, False)
        'Call UpdateEnergy(-5, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherWater_Click()
    If Energy >= 5 Then
        WaterGather
        Call UpdateActionsCompleted(1, False)
        'Call UpdateEnergy(-5, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub GatherStones_Click()
    If Energy >= 5 Then
        StoneGather
        Call UpdateActionsCompleted(1, False)
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

Private Sub KnifeCraft_Click()
    CraftKnife
    Call UpdateActionsCompleted(1, False)
End Sub

Private Sub KnifeCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Craft a knife using a stick and a stone. The knife allows you to skin rabbits and craft more."
End Sub

Private Sub KnifeLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: The knife allows you to skin animals and can be used to increase crafting outputs."
End Sub

Private Sub MeatConsume_Click()
    If Meat > 0 Then
        ConsumeMeat
        NewsLabel.Caption = "News: You ate some meat."
    Else
        NewsLabel.Caption = "News: You don't have any meat."
    End If
End Sub

Private Sub RawMeatConsume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Raw meat has a chance to decrease your health and provides less increase in stats when uncooked."
End Sub

Private Sub RawMeatCook_Click()
    If RawMeat > 0 Then
        CookRawMeat
        Call UpdateThingsCooked(1, False)
    Else
        NewsLabel.Caption = "News: You don't have meat to cook."
    End If
End Sub

Private Sub RabbitHunt_Click()
    If Energy >= 6 Then
        HuntRabbit
        Call UpdateActionsCompleted(1, False)
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub RawMeatConsume_Click()
    If RawMeat > 0 Then
        ConsumeRawMeat
    End If
End Sub

Private Sub SnipeHunt_Click()
    If SnipeEnabled = False Then
        Call UpdateEnergy(-10, False)
        NewsLabel.Caption = "News: You failed to catch a snipe."
    Else
        Call UpdateMeat(999, False)
    End If
End Sub

Private Sub SpearCraft_Click()
    If Spear = False Then
        CraftSpear
        Call UpdateActionsCompleted(1, False)
    End If
End Sub

Private Sub FireCraft_Click()
    If Fire = False Then
        CraftFire
        Call UpdateActionsCompleted(1, False)
    End If
End Sub

Private Sub SSBuild_Click()
    If StickShelt = False Then
        BuildSShelter
        Call UpdateActionsCompleted(1, False)
    End If
End Sub

Private Sub StickCraft_Click()
    CraftSticks
    Call UpdateActionsCompleted(1, False)
End Sub

Private Sub StickCraft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Turn a log into sticks! Requires 1 log."
End Sub

Private Sub TrapBuild_Click()
    BuildTrap
    Call UpdateActionsCompleted(1, False)
End Sub

Private Sub TrapBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Use 2 stones, 5 sticks, and 2 berries to build a trap for catching rabbits."
End Sub

Private Sub TrapLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Traps have a chance to catch rabbits if they are set. You have " & SetTraps & " traps set."
End Sub

Private Sub TrapMaintain_Click()
    MaintainTraps
End Sub

Private Sub TrapMaintain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'SetTraps = 0
    'For i = 1 To Traps
        'If TrapsSet(i) = True Then SetTraps = SetTraps + 1
    'Next i
    InfoLabel.Caption = "Info: Use two berries to reset a trap."
End Sub

Private Sub VineGather_Click()
    If Energy >= 5 Then
        Call GatherVines
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub VineGather_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click this button to gather vines. Costs 5 energy."
End Sub

Private Sub WaterConsume_Click()
    ConsumeWater
End Sub

Private Sub HealthLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: If your health reaches 0, you die. Your health increases slowly over time."
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
    InfoLabel.Caption = "Info: Click this button to make a spear. You need 2 sticks, and 2 stones."
End Sub

Private Sub FireLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FireStartDay > 0 Then
        InfoLabel.Caption = "Info: A fire allows you to cook food. Your fire has " & 5 - (DayNum - FireStartDay) & " days until it goes out."
    Else
        InfoLabel.Caption = "Info: A fire allows you to cook food. Visit the building tab for more information."
    End If
End Sub

Private Sub SSBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Click to build a stick shelter. It requires 15 sticks and 50 leaves."
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

Private Sub RawMeatLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Uncooked meat can be eaten for significantly reduced stat increase and a chance to lose health."
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

Private Sub WolfFeedButton_Click()
    WolfFeed
End Sub

Private Sub WolfFeedButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Feed your wolf some raw or cooked meet to maintain his health and hunger."
End Sub

Private Sub WolfHunt_Click()
    If Energy >= 9 Then
        HuntWolves
    Else
        NotEnoughEnergy
    End If
End Sub

Private Sub WolfHuntButton_Click()
    If CWolfDailyAction = False Then
        CWolfHunt
    Else
        NewsLabel.Caption = "News: Your wolf has already hunted once today."
    End If
End Sub

Private Sub WolfHuntButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Command your wolf to hunt. He can only do it once a day!"
End Sub

Private Sub WolfLabel_Click()
    If CWolf = True Then
        CompActFrame.Caption = "Wolf - Actions"
        CompStatFrame.Caption = "Wolf - Stats"
    Else
        NewsLabel.Caption = "News: You don't have the wolf companion yet."
    End If
End Sub

Private Sub WolfLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: This is your wolf companion. Don't forget to feed him!"
End Sub

Private Sub WolfPetButton_Click()
    WolfPet
End Sub

Private Sub WolfPetButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InfoLabel.Caption = "Info: Spend 3 energy to bond with your wolf."
End Sub
