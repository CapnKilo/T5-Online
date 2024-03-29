Attribute VB_Name = "modGeneral"
Option Explicit

' Pauses the thread of execution
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

' Timer functions
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' For clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

' DirectDraw master engine for early binding
Public DX7 As New DirectX7

Public Sub Main()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set loading screen
    loadGUI True
    frmLoad.Visible = True

    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions

    ' load main menu
    Call SetStatus("Loading Menu...")
    Load frmMenu

    ' load music/sound engine
    InitSound
    
    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)
    
    ' load gui
    Call SetStatus("Loading interface...")
    loadGUI
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "Skillicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\", "main"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "bars"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    
    ' clear old data
    Call SetStatus("Clearing game data...")
    Call ClearGameData
    
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34)
    
    ' update the form's caption with the game's name
    frmMain.Caption = Options.Game_Name
    
    ' initialize DirectX
    If Not InitDirectDraw Then
        MsgBox "Error Initializing DirectX7 - DirectDraw."
        DestroyGame
    End If
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("Initializing DirectX...")
    
    ' DX7 Master Object is already created, early binding
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSkillIcons
    Call CheckFaces
    
    ' Reset values
    Ping = -1

    ' cache the buttons then reset & render them
    Call SetStatus("Loading buttons...")
    cacheButtons
    resetButtons_Menu
    
    ' we can now see it
    frmMenu.Visible = True
    
    ' hide all pics
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    
    ' hide the load form
    frmLoad.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub loadGUI(Optional ByVal loadingScreen As Boolean = False)
Dim i As Long

    ' if we can't find the interface
    On Error GoTo errorhandler
    
    ' loading screen
    If loadingScreen Then
        frmLoad.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\loading.jpg")
        Exit Sub
    End If

    ' menu
    frmMenu.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\background.jpg")
    frmMenu.picMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\main.jpg")
    frmMenu.picLogin.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\login.jpg")
    frmMenu.picRegister.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\register.jpg")
    frmMenu.picCredits.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\credits.jpg")
    frmMenu.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\character.jpg")
    ' main
    frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main.jpg")
    frmMain.picInventory.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\inventory.jpg")
    frmMain.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\character.jpg")
    frmMain.picSkills.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\Skills.jpg")
    frmMain.picOptions.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\options.jpg")
    frmMain.picParty.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\party.jpg")
    frmMain.picItemDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_item.jpg")
    frmMain.picSkillDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_Skill.jpg")
    frmMain.picTempInv.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempskill.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picShop.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\shop.jpg")
    frmMain.picBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bank.jpg")
    frmMain.picTrade.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\trade.jpg")
    frmMain.picHotbar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\hotbar.jpg")
    frmMain.picQuests.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\quests.jpg")
    ' main - bars
    frmMain.imgHPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\health.jpg")
    frmMain.imgMPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\spirit.jpg")
    frmMain.imgEXPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.jpg")
    ' main - party bars
    For i = 1 To MAX_PARTY_MEMBERS
        frmMain.imgPartyHealth(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_health.jpg")
        frmMain.imgPartySpirit(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_spirit.jpg")
    Next
    
    ' store the bar widths for calculations
    HPBar_Width = frmMain.imgHPBar.width
    SPRBar_Width = frmMain.imgMPBar.width
    EXPBar_Width = frmMain.imgEXPBar.width
    ' party
    Party_HPWidth = frmMain.imgPartyHealth(1).width
    Party_SPRWidth = frmMain.imgPartySpirit(1).width
    
    Exit Sub
    
' let them know we can't load the GUI
errorhandler:
    MsgBox "Cannot find one or more interface images." & vbNewLine & "If they exist then you have not extracted the project properly." & vbNewLine & "Please follow the installation instructions fully.", vbCritical
    DestroyGame
    Exit Sub
End Sub

Public Sub MenuState(ByVal state As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmLoad.Visible = True

    Select Case state
        Case MENU_STATE_ADDCHAR
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending character addition data...")

                If frmMenu.optMale.value Then
                    Call SendAddChar(frmMenu.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                Else
                    Call SendAddChar(frmMenu.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                End If
            End If
            
        Case MENU_STATE_NEWACCOUNT
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMenu.txtRUser.Text, frmMenu.txtRPass.Text)
            End If

        Case MENU_STATE_LOGIN
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMenu.txtLUser.Text, frmMenu.txtLPass.Text)
                Exit Sub
            End If
    End Select

    If frmLoad.Visible Then
        If Not IsConnected Then
            frmMenu.Visible = True
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False
            frmMenu.picMain.Visible = True
            frmLoad.Visible = False
            Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, Options.Game_Name)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim buffer As clsBuffer, i As Long

    isLogging = True
    InGame = False
    Set buffer = New clsBuffer
    buffer.WriteLong CQuit
    SendData buffer.ToArray()
    Set buffer = Nothing
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    InvX = 0
    InvY = 0
    EqX = 0
    EqY = 0
    SkillX = 0
    SkillY = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SkillBuffer = 0
    SkillBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Skill
    Unload frmEditor_Conv
    
    ' hide main form stuffs
    frmMenu.picMain.Visible = True
    frmMain.txtChat.Text = vbNullString
    frmMain.txtMyChat.Text = vbNullString
    frmMain.picCurrency.Visible = False
    frmMain.picDialogue.Visible = False
    frmMain.picInventory.Visible = True
    frmMain.picTrade.Visible = False
    frmMain.picCover.Visible = False
    frmMain.picSkills.Visible = False
    frmMain.picCharacter.Visible = False
    frmMain.picOptions.Visible = False
    frmMain.picParty.Visible = False
    frmMain.picBank.Visible = False
    frmMain.picConv.Visible = False
End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EnteringGame = True
    frmMenu.Visible = False
    EnteringGame = False
    
    ' bring all the main gui components to the front
    frmMain.picShop.ZOrder (0)
    frmMain.picBank.ZOrder (0)
    frmMain.picTrade.ZOrder (0)
    
    ' hide gui
    frmMain.picCover.Visible = False
    InBank = False
    InShop = False
    InTrade = False
    
    ' show the main form
    frmLoad.Visible = False
    frmMain.Show
    
    ' Set the focus
    frmMain.picScreen.Visible = True
    Call SetFocusOnChat
    frmMain.txtMyChat.Visible = False
    
    ' blt inv
    BltInventory
    
    ' blt hotbar
    BltHotbar
    
    ' get ping
    GetPing
    DrawPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.max = MAX_ITEMS
    frmMain.scrlAItem.value = 1
    
    ' stop the song playing
    StopMusic
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' break out of GameLoop
    InGame = False
    Call DestroyTCP
    
    ' destroy objects in reverse order
    Call DestroyDirectDraw

    ' destory DirectX7 master object
    If Not DX7 Is Nothing Then
        Set DX7 = Nothing
    End If

    Call UnloadAllForms
    
    ' unload the music engine
    CloseSound
    
    ' close the application
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmLoad.lblStatus.Caption = Caption
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetFocusOnChat()

    On Error Resume Next ' prevent RTE5, no way to handle error
    
    If frmMain.txtMyChat.Visible = True Then
        frmMain.txtMyChat.SetFocus
    Else
        frmMain.txtMyChat.Visible = True
        frmMain.txtMyChat.SetFocus
        frmMain.txtMyChat.Visible = False
    End If
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GlobalX As Long
Dim GlobalY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GlobalX = PB.Left
    GlobalY = PB.top

    If Button = 1 Then
        PB.Left = GlobalX + X - SOffsetX
        PB.top = GlobalY + Y - SOffsetY
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MovePicture", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For i = 1 To Len(sInput)
        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If
    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' menu - login
    With MenuButton(1)
        .FileName = "login"
        .state = 0 ' normal
    End With
    
    ' menu - register
    With MenuButton(2)
        .FileName = "register"
        .state = 0 ' normal
    End With
    
    ' menu - credits
    With MenuButton(3)
        .FileName = "credits"
        .state = 0 ' normal
    End With
    
    ' menu - exit
    With MenuButton(4)
        .FileName = "exit"
        .state = 0 ' normal
    End With
    
    ' main - inv
    With MainButton(1)
        .FileName = "inv"
        .state = 0 ' normal
    End With
    
    ' main - Skills
    With MainButton(2)
        .FileName = "Skills"
        .state = 0 ' normal
    End With
    
    ' main - char
    With MainButton(3)
        .FileName = "char"
        .state = 0 ' normal
    End With
    
    ' main - opt
    With MainButton(4)
        .FileName = "opt"
        .state = 0 ' normal
    End With
    
    ' main - trade
    With MainButton(5)
        .FileName = "trade"
        .state = 0 ' normal
    End With
    
    ' main - party
    With MainButton(6)
        .FileName = "party"
        .state = 0 ' normal
    End With
    
    ' main - quests
    With MainButton(7)
        .FileName = "quests"
        .state = 0 ' normal
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MenuButton(i).state = 0 'normal
            renderButton_Menu i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Menu = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Menu(ByVal buttonNum As Long)
Dim bSuffix As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MenuButton(buttonNum).state
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMenu.imgButton(buttonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(buttonNum).FileName & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Menu(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MenuButton(buttonNum).state = bState
        renderButton_Menu buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' main specific buttons
Public Sub resetButtons_Main(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MAINBUTTONS
        ' only change if different and not exception
        If Not MainButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MainButton(i).state = 0 'normal
            renderButton_Main i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Main = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Main(ByVal buttonNum As Long)
Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MainButton(buttonNum).state
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMain.imgButton(buttonNum).Picture = LoadPicture(App.Path & MAINBUTTON_PATH & MainButton(buttonNum).FileName & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Main(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MainButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MainButton(buttonNum).state = bState
        renderButton_Main buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
Dim strLoad As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = Dir(App.Path & MUSIC_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    ' Cache sound list
    strLoad = Dir(App.Path & SOUND_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearGameData()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ClearNPCs
    Call ClearResources
    Call ClearItems
    Call ClearShops
    Call ClearSkills
    Call ClearAnimations

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearGameData", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
