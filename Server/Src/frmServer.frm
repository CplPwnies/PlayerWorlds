VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmServer 
   Caption         =   "Playerworlds Server"
   ClientHeight    =   3375
   ClientLeft      =   4965
   ClientTop       =   4335
   ClientWidth     =   7695
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5953
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Chat"
      TabPicture(0)   =   "frmServer.frx":030C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtChat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtText"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Bug Report"
      TabPicture(1)   =   "frmServer.frx":0328
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstBugReport"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Game Information"
      TabPicture(2)   =   "frmServer.frx":0344
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(4)=   "Frame1"
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame1 
         Caption         =   "Game IP"
         Height          =   615
         Left            =   -69720
         TabIndex        =   13
         Top             =   360
         Width           =   2055
         Begin VB.TextBox txtIP 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Game Port"
         Height          =   615
         Left            =   -69960
         TabIndex        =   11
         Top             =   960
         Width           =   1215
         Begin VB.TextBox txtPort 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Total Online"
         Height          =   615
         Left            =   -68640
         TabIndex        =   9
         Top             =   960
         Width           =   1215
         Begin VB.TextBox txtOnline 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Accounts Online"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   2295
         Begin VB.ListBox lstAccounts 
            Height          =   2400
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Players Online"
         Height          =   2775
         Left            =   -72360
         TabIndex        =   5
         Top             =   480
         Width           =   2295
         Begin VB.ListBox lstPlayers 
            Height          =   2400
            ItemData        =   "frmServer.frx":0360
            Left            =   120
            List            =   "frmServer.frx":0362
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.ListBox lstBugReport 
         Height          =   2790
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   7455
      End
      Begin VB.TextBox txtText 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   7455
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   7455
      End
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   240
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   240
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   10000
      Left            =   240
      Top             =   240
   End
   Begin VB.Timer tmrReboot 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   240
   End
   Begin VB.Timer PlayerTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2640
      Top             =   240
   End
   Begin VB.Timer tmrtTime 
      Interval        =   1000
      Left            =   3120
      Top             =   240
   End
   Begin MSWinsockLib.Winsock SocketIP 
      Left            =   7200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrWarpPlayer 
      Interval        =   25
      Left            =   3600
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bug Reports"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuShutdown 
         Caption         =   "&Shutdown"
      End
      Begin VB.Menu mnuServerReboot 
         Caption         =   "&Reboot"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuSetAccess 
         Caption         =   "Set &Access"
      End
      Begin VB.Menu mnuReloadClasses 
         Caption         =   "Reload &Classes"
      End
      Begin VB.Menu mnuReloadScripts 
         Caption         =   "Reload &Scripts"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server L&og"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPlayers 
      Caption         =   "&Players"
      Visible         =   0   'False
      Begin VB.Menu mnuWarn 
         Caption         =   "&Warn"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "&Mute"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "&Kick"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "&Ban"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim lmsg As Long

    lmsg = X / Screen.TwipsPerPixelX
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub

Private Sub Form_Resize()
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If
End Sub

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub lstPlayers_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbKeyRButton Then
        Me.PopupMenu mnuPlayers
    End If
End Sub

Private Sub mnuBan_Click()
    Dim N As Long
    N = FindPlayer(lstPlayers.List(lstPlayers.ListIndex))

    If N > 0 Then
        Call BanIndex(N, "Server")
        Call GlobalMsg(GetPlayerName(N) & " has been banned from " & GAME_NAME & " by the Server!", White)
        Call AddLog(GetPlayerName(N) & " has banned by the Server.", ADMIN_LOG)
        Call AlertMsg(N, "You have been banned by the Server!")
    End If

End Sub

Private Sub mnuKick_Click()
    Dim N As Long
    N = FindPlayer(lstPlayers.List(lstPlayers.ListIndex))

    If N > 0 Then
        Call GlobalMsg(GetPlayerName(N) & " has been kicked from " & GAME_NAME & " by the Server!", White)
        Call AddLog("The Server has kicked " & GetPlayerName(N) & ".", ADMIN_LOG)
        Call AlertMsg(N, "You have been kicked by the server!")
    End If
End Sub

Private Sub mnuReloadScripts_Click()
    Call ReloadScripts
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.Checked = True Then
        mnuServerLog.Checked = False
        ServerLog = False
    Else
        mnuServerLog.Checked = True
        ServerLog = True
    End If
End Sub

Private Sub mnuServerReboot_Click()
    Me.tmrReboot.Enabled = True
End Sub

Private Sub mnuSetAccess_Click()
    Dim name As String
    Dim I As Integer, PlayerAccess As Byte

    name = InputBox("What is the ONLINE player's name?", "Give Access to whom?", "")

    If name <> vbNullString Then
        I = FindPlayer(name)
    Else
        MsgBox ("Player Not Online!")
        Exit Sub
    End If

    PlayerAccess = InputBox("What access level?", "Access (1-6):", "1")

    If IsConnected(I) Then
        If PlayerAccess <> 0 Then
            ' sloppy... but whatever
            Call SetPlayerAccess(I, PlayerAccess)

            Call PlayerMsg(I, "Your access has been changed.", BrightRed)
            Call SendPlayerData(I)
        Else
            MsgBox ("Invalid Access!")
        End If
    Else
        MsgBox ("Player Not Online!")
    End If
End Sub

Private Sub mnuWarn_Click()
    Dim N As Long
    Dim Msg As String
    N = FindPlayer(lstPlayers.List(lstPlayers.ListIndex))

    If N > 0 Then
        Msg = InputBox("What is the warning?", GAME_NAME, "")
        Call PlayerMsg(N, "Server Warning: " & Msg, BrightRed)
        Call AddLog("The Server has warned " & GetPlayerName(N) & ".", ADMIN_LOG)
    End If
End Sub

Private Sub PlayerTimer_Timer()
    Dim I As Long


    If PlayerI <= HighIndex Then
        If IsPlaying(PlayerI) Then
            Call SavePlayer(PlayerI)
            Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & ", you have been saved.", BrightGreen)
        End If
        PlayerI = PlayerI + 1

    ElseIf PlayerI > HighIndex Then
        PlayerI = 1
        PlayerTimer.Enabled = False
        tmrPlayerSave.Enabled = True
    End If
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub tmrtTime_Timer()
    Static I As Long

    I = I + 1
    If I Mod 31536000 = 0 Then
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 6 & "," & I
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 4 & "," & I
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 3 & "," & I
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 2 & "," & I
        Exit Sub
    ElseIf I Mod 604800 = 0 Then
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 4 & "," & I
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 3 & "," & I
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 2 & "," & I
        Exit Sub
    ElseIf I Mod 86400 = 0 Then
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 3 & "," & I
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 2 & "," & I
        Exit Sub
    ElseIf I Mod 3600 = 0 Then
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 2 & "," & I
        Exit Sub
    ElseIf I Mod 60 = 0 Then
        MyScript.ExecuteStatement "\scripts\Main.as", "OnTime " & 1 & "," & I
        Exit Sub
    End If

End Sub

Private Sub tmrWarpPlayer_Timer()
    Call CheckWarp
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim$(txtChat.Text) <> vbNullString Then
        Call GlobalMsg(txtChat.Text, White)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = vbNullString
    End If
End Sub

Private Sub tmrShutdown_Timer()
    Static Secs As Long

    If ShutOn = False Then
        Secs = 30
        Call TextAdd(frmServer.txtText, "Automated Server Shutdown Canceled!", True)
        Call GlobalMsg("Server Shutdown Canceled!", BrightBlue)
        tmrShutdown.Enabled = False
        Exit Sub
    End If
    If Secs <= 0 Then Secs = 30
    Secs = Secs - 1
    Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", Yellow)
    Call TextAdd(frmServer.txtText, "Automated Server Shutdown in " & Secs & " seconds.", True)
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

Private Sub tmrReboot_Timer()
    Static Secs As Long

    If Secs <= 0 Then Secs = 30
    Secs = Secs - 2
    Call GlobalMsg("Server Reboot in " & Secs & " seconds.", Yellow)
    Call TextAdd(frmServer.txtText, "Automated Server Reboot in " & Secs & " seconds.", True)
    If Secs <= 0 Then
        tmrReboot.Enabled = False
        Shell (App.Path & "/Bootstrap.exe"), vbNormalFocus
        Call DestroyServer
    End If
End Sub

Private Sub mnuShutdown_Click()
    If ShutOn = False Then
        tmrShutdown.Enabled = True
        mnuShutdown.Caption = "Cancel Shutdown"
        ShutOn = True
    ElseIf ShutOn = True Then
        mnuShutdown.Caption = "Shutdown"
        ShutOn = False
    End If
End Sub

Private Sub Form_Load()
    ShutOn = False
    ServerLog = True
    Me.Caption = GAME_NAME & " :: Server"
    Me.txtIP.Text = SocketIP.LocalIP
End Sub

Private Sub mnuExit_Click()
    Call DestroyServer
End Sub

Private Sub mnuReloadClasses_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText, "All classes reloaded.", True)
End Sub

' Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
' Call AcceptConnection(index, requestID)
' End Sub

' Private Sub Socket_Accept(index As Integer, SocketId As Integer)
' Call AcceptConnection(index, SocketId)
' End Sub

' Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
' If IsConnected(index) Then
' Call IncomingData(index, bytesTotal)
' End If
' End Sub

' Private Sub Socket_Close(index As Integer)
' Call CloseSocket(index)
' End Sub

