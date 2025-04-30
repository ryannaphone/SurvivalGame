VERSION 5.00
Begin VB.Form RenameForm 
   Caption         =   "RenameCompanions"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EntertButton 
      Caption         =   "Rename"
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
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Frame CompanionFrame 
      Caption         =   "Select a companion"
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
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.Label WolfLabel 
         Caption         =   "Worf the Wolf"
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
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.TextBox NameText 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Enter companion name below."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
   End
End
Attribute VB_Name = "RenameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CWolfSelected As Boolean

Private Sub EntertButton_Click()
    If CWolfSelected = True Then
        If NameText.Text <> "" Then
            CWolfName = NameText.Text
            MainForm.WolfLabel.Caption = CWolfName
            Unload RenameForm
        Else
            MsgBox "Please enter a name first."
        End If
    Else
        MsgBox "Please select a companion to rename first."
    End If
End Sub

Private Sub Form_Load()
    If CWolf = True Then
        CWolfSelected = False
        WolfLabel.Visible = True
    Else
        WolfLabel.Visible = False
        CWolfSelected = False
    End If
End Sub

Private Sub WolfLabel_Click()
    CWolfSelected = True
    WolfLabel.ForeColor = &H8000000D
End Sub
