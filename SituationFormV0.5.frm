VERSION 5.00
Begin VB.Form SituationForm 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SituationsPrompt"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Option2 
      Caption         =   "Second Option"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   4215
   End
   Begin VB.CommandButton Option1 
      Caption         =   "First Option"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label SituationInfo 
      BackColor       =   &H00000000&
      Caption         =   "Situation Info: This is gonna show information about your current situation. Maybe there is a wolf you can tame ;)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "SituationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
    Select Case Option1.Caption
        Case "Try to kill them both."
            Call UpdateWolfBond(-1, False)
            Option1.Enabled = False
            Option2.Enabled = False
            SituationInfo.Caption = "Situation Info: You charge them, but both escape."
            WaitForTime (2.5)
            Unload SituationForm
        Case "Offer it some meat. (Requires 1 raw or cooked meat)"
            If RawMeat >= 1 Then
                Call UpdateWolfBond(1, False)
                Call UpdateRawMeat(-1, False)
                Option1.Enabled = False
                Option2.Enabled = False
                SituationInfo.Caption = "Situation Info: You offer the wolf some meat and it cautiously accepts before running off."
                WaitForTime (2.5)
                Unload SituationForm
            ElseIf Meat >= 1 Then
                Call UpdateWolfBond(1, False)
                Call UpdateMeat(-1, False)
                Option1.Enabled = False
                Option2.Enabled = False
                SituationInfo.Caption = "Situation Info: You offer the wolf some meat and it cautiously accepts before running off."
                WaitForTime (2.5)
                Unload SituationForm
            Else
                Call UpdateWolfBond(0, False)
                Option1.Enabled = False
                Option2.Enabled = False
                SituationInfo.Caption = "Situation Info: You don't have any meat to offer and the wolf runs off."
                WaitForTime (2.5)
                Unload SituationForm
            End If
        Case "Attack it."
            Call UpdateWolfBond(-1, False)
            Option1.Enabled = False
            Option2.Enabled = False
            SituationInfo.Caption = "Situation Info: You attack the wolf and hit it once before it runs away."
            WaitForTime (2.5)
            Unload SituationForm
    End Select
End Sub

Private Sub Option2_Click()
    Select Case Option2.Caption
        Case "Leave them alone."
            Call UpdateWolfBond(1, False)
            Option1.Enabled = False
            Option2.Enabled = False
            SituationInfo.Caption = "Situation Info: You turn around and walk away."
            WaitForTime (2.5)
            Unload SituationForm
        Case "Throw a rock at it."
            Call UpdateWolfBond(-1, False)
            Option1.Enabled = False
            Option2.Enabled = False
            SituationInfo.Caption = "Situation Info: The rock lands near the wolf and it scampers off."
            WaitForTime (2.5)
            Unload SituationForm
        Case "Offer it some meat. (Requires 1 raw or cooked meat)"
            If RawMeat >= 1 Then
                Call UpdateWolfBond(1, False)
                Call UpdateRawMeat(-1, False)
                Option1.Enabled = False
                Option2.Enabled = False
                SituationInfo.Caption = "Situation Info: You offer the wolf some meat and it cautiously accepts before running off."
                WaitForTime (2.5)
                Unload SituationForm
            ElseIf Meat >= 1 Then
                Call UpdateWolfBond(1, False)
                Call UpdateMeat(-1, False)
                Option1.Enabled = False
                Option2.Enabled = False
                SituationInfo.Caption = "Situation Info: You offer the wolf some meat and it cautiously accepts before running off."
                WaitForTime (2.5)
                Unload SituationForm
            Else
                Call UpdateWolfBond(0, False)
                Option1.Enabled = False
                Option2.Enabled = False
                SituationInfo.Caption = "Situation Info: You don't have any meat to offer and the wolf runs off."
                WaitForTime (2.5)
                Unload SituationForm
            End If
    End Select
End Sub

