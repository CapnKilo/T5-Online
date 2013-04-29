VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Editor"
   ClientHeight    =   8190
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   18
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstIndex 
         Height          =   7470
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   855
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tasks"
      Height          =   5775
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   5175
      Begin VB.HScrollBar scrlXP 
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   5040
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   840
         Min             =   1
         TabIndex        =   22
         Top             =   5400
         Value           =   1
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Value           =   1
         Width           =   2655
      End
      Begin VB.HScrollBar scrlTaskCount 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   10
         Top             =   480
         Value           =   1
         Width           =   4935
      End
      Begin VB.Frame fraTaskData 
         Caption         =   "Task 1/50"
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4935
         Begin VB.TextBox txtInfo 
            Height          =   1335
            Left            =   120
            TabIndex        =   24
            Top             =   1680
            Width           =   4695
         End
         Begin VB.HScrollBar scrlTask 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   6
            Top             =   240
            Value           =   1
            Width           =   4695
         End
         Begin VB.ComboBox cmbTaskType 
            Height          =   315
            ItemData        =   "frmEditor_Quest.frx":0000
            Left            =   120
            List            =   "frmEditor_Quest.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   4695
         End
         Begin VB.HScrollBar scrlDataIndex 
            Height          =   255
            Left            =   2040
            TabIndex        =   4
            Top             =   3120
            Width           =   2775
         End
         Begin VB.HScrollBar scrlDataValue 
            Height          =   255
            Left            =   2040
            TabIndex        =   3
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label Label3 
            Caption         =   "Task Description"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   4695
         End
         Begin VB.Label Label2 
            Caption         =   "TaskType"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   4695
         End
         Begin VB.Label lblDataIndex 
            Caption         =   "Data Index: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3120
            Width           =   2055
         End
         Begin VB.Label lblDataValue 
            Caption         =   "Data Value: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   3480
            Width           =   2055
         End
      End
      Begin VB.Label lblXP 
         Caption         =   "XP Reward: None"
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   5400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblReward 
         Caption         =   "Reward: None"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label lblTaskCount 
         Caption         =   "Task Count: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   7560
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTaskType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Quest(lstIndex.ListIndex + 1).Task(scrlTask.value).TaskType = cmbTaskType.ListIndex
    
    scrlDataIndex.Visible = False
    scrlDataValue.Visible = False
    lblDataIndex.Visible = False
    lblDataValue.Visible = False
    
    Select Case cmbTaskType.ListIndex
        Case 2 ' kill
            scrlDataIndex.Visible = True
            scrlDataValue.Visible = True
            lblDataIndex.Visible = True
            lblDataValue.Visible = True
            lblDataIndex.Caption = "NPC Index: " & scrlDataIndex.value
            lblDataValue.Caption = "NPC Value: " & scrlDataValue.value
        Case 4 ' resource
            scrlDataIndex.Visible = True
            scrlDataValue.Visible = True
            lblDataIndex.Visible = True
            lblDataValue.Visible = True
            lblDataIndex.Caption = "Resource Index: " & scrlDataIndex.value
            lblDataValue.Caption = "Resource Value: " & scrlDataValue.value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTaskType_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call QuestEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call QuestEditorOK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlTaskCount.max = MAX_QUEST_TASKS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Re-init the editor
    If QuestEditorLoaded = True Then QuestEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblReward.Caption = "Reward: " & scrlAmount.value & "x " & Trim$(Item(scrlReward.value).name)
    Quest(lstIndex.ListIndex + 1).RewardAmount = scrlAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAmount_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDataIndex_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblDataIndex.Caption = "Data Index: " & scrlDataIndex.value
    Quest(lstIndex.ListIndex + 1).Task(scrlTask.value).DataIndex = scrlDataIndex.value

    Select Case cmbTaskType.ListIndex
        Case 2 ' kill
            lblDataIndex.Caption = "NPC Index: " & scrlDataIndex.value
            lblDataValue.Caption = "NPC Value: " & scrlDataValue.value
        Case 4 ' resource
            lblDataIndex.Caption = "Resource Index: " & scrlDataIndex.value
            lblDataValue.Caption = "Resource Value: " & scrlDataValue.value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDataIndex_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDataValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblDataValue.Caption = "Data Value: " & scrlDataValue.value
    Quest(lstIndex.ListIndex + 1).Task(scrlTask.value).DataAmount = scrlDataValue.value
    
    Select Case cmbTaskType.ListIndex
        Case 2 ' kill
            lblDataIndex.Caption = "NPC Index: " & scrlDataIndex.value
            lblDataValue.Caption = "NPC Value: " & scrlDataValue.value
        Case 4 ' resource
            lblDataIndex.Caption = "Resource Index: " & scrlDataIndex.value
            lblDataValue.Caption = "Resource Value: " & scrlDataValue.value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDataValue_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlReward.value <> 0 Then
        If Item(scrlReward.value).Type = ITEM_TYPE_CURRENCY Then
            scrlAmount.Visible = True
            lblAmount.Visible = True
            lblReward.Caption = "Reward: " & scrlAmount.value & "x " & Trim$(Item(scrlReward.value).name)
        Else
            lblReward.Caption = "Reward: " & Trim$(Item(scrlReward.value).name)
            Quest(lstIndex.ListIndex + 1).RewardAmount = 1
            scrlAmount.Visible = False
            lblAmount.Visible = False
        End If
    Else
        lblReward.Caption = "Reward: None"
    End If
    Quest(lstIndex.ListIndex + 1).Reward = scrlReward.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlReward_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTask_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If QuestEditorLoaded = False Then Exit Sub
    fraTaskData.Caption = "Task " & scrlTask.value & "/" & scrlTaskCount.value
    
    scrlDataValue.value = Quest(EditorIndex).Task(scrlTask.value).DataAmount
    scrlDataIndex.value = Quest(EditorIndex).Task(scrlTask.value).DataIndex
    
    lblDataValue.Caption = "Data Value: " & scrlDataValue.value
    lblDataIndex.Caption = "Data Index: " & scrlDataIndex.value
    cmbTaskType.ListIndex = Quest(EditorIndex).Task(scrlTask.value).TaskType
    txtInfo.Text = Trim$(Quest(EditorIndex).Task(scrlTask.value).info)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTask_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTaskCount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblTaskCount.Caption = "Task Count: " & scrlTaskCount.value
    Quest(EditorIndex).TaskCount = scrlTaskCount.value
    
    If scrlTaskCount.value < scrlTask.value Then scrlTask.value = scrlTaskCount.value
    scrlTask.max = scrlTaskCount.value
    fraTaskData.Caption = "Task " & scrlTask.value & "/" & scrlTaskCount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTaskCount_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlXP.value > 0 Then
        lblXP.Caption = "XP Reward: " & scrlXP.value
    Else
        lblXP.Caption = "XP Reward: None"
    End If
    
    Quest(EditorIndex).XPReward = scrlXP.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlXP_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtInfo_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Quest(EditorIndex).Task(scrlTask.value).info = txtInfo.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtInfo_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Quest(EditorIndex).name = txtName.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

