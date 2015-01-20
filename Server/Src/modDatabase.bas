Attribute VB_Name = "modDatabase"
Option Explicit

Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String          ' Max string length
    Dim szReturn As String         ' Return default value if not found

    szReturn = vbNullString

    sSpaces = Space$(5000)

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
' ****************************************************************
' * WHEN    WHO    WHAT
' * ----    ---    ----
' * 07/16/2005  Shannara   Optimized function.
' ****************************************************************

    If RAW = False Then
        If Dir(App.Path & "\" & FileName) = vbNullString Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If Dir(FileName) = vbNullString Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
        End If
    End If
End Function

Sub SavePlayer(ByVal Index As Long)
    Dim FileName As String
    Dim f As Long
    Dim I As Long
    Dim StartByte As Long

    FileName = App.Path & "\data\accounts\" & Trim$(Player(Index).Login) & ".act"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Player(Index).Login
    Put #f, , Player(Index).Password

    Put #f, , Player(Index).Char(1)
    Put #f, , Player(Index).Char(2)
    Put #f, , Player(Index).Char(3)
    Close #f
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim FileName As String
    Dim f As Long
    Dim StartByte As Long

    Call ClearPlayer(Index)

    FileName = App.Path & "\data\accounts\" & Trim$(Name) & ".act"

    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , Player(Index).Login
    Get #f, , Player(Index).Password

    Get #f, , Player(Index).Char(1)
    Get #f, , Player(Index).Char(2)
    Get #f, , Player(Index).Char(3)
    Close #f
End Sub

Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String

    FileName = "\data\accounts\" & Trim$(Name) & ".act"

    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim$(Player(Index).Char(CharNum).Name) <> vbNullString Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim FileName As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Integer

    nFileNum = FreeFile

    PasswordOK = False

    If AccountExist(Name) Then
        FileName = App.Path & "\data\accounts\" & Trim$(Name) & ".act"

        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
        Get #nFileNum, NAME_LENGTH, RightPassword
        Close #nFileNum

        If Trim$(Password) = Trim$(RightPassword) Then
            PasswordOK = True
        End If
    End If


End Function

Function EncKeyOK(ByVal Name As String, ByVal EncKey As String) As Boolean
    Dim FileName As String
    Dim RightPassword As String
    Dim RightEncKey As String

    EncKeyOK = False

    If AccountExist(Name) Then
        FileName = App.Path & "\data\accounts\" & Trim$(Name) & ".act"
        RightEncKey = ENC_KEY

        If UCase$(Trim$(EncKey)) = UCase$(Trim$(ENC_KEY)) Then
            EncKeyOK = True
        End If
    End If
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim I As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    Player(Index).EncKey = ENC_KEY

    For I = 1 To MAX_CHARS
        Call ClearChar(Index, I)
    Next I

    Call SavePlayer(Index)
End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
    Dim f As Long

    If Trim$(Player(Index).Char(CharNum).Name) = vbNullString Then
        Player(Index).CharNum = CharNum

        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = ClassNum

        If Player(Index).Char(CharNum).Sex = SEX_MALE Then
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        Else
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).FSprite
        End If

        Player(Index).Char(CharNum).Level = 1

        Player(Index).Char(CharNum).STR = Class(ClassNum).STR
        Player(Index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(Index).Char(CharNum).SPEED = Class(ClassNum).SPEED
        Player(Index).Char(CharNum).MAGI = Class(ClassNum).MAGI

        Player(Index).Char(CharNum).Map = START_MAP
        Player(Index).Char(CharNum).X = START_X
        Player(Index).Char(CharNum).y = START_Y

        Player(Index).Char(CharNum).HP = GetPlayerMaxHP(Index)
        Player(Index).Char(CharNum).MP = GetPlayerMaxMP(Index)
        Player(Index).Char(CharNum).SP = GetPlayerMaxSP(Index)

        ' Append name to file
        f = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #f
        Print #f, Name
        Close #f

' Call SavePlayer(Index)

        Exit Sub
    End If
End Sub

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
    Dim f1 As Long, f2 As Long
    Dim s As String

    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String

    FindChar = False

    f = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #f
    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If
    Loop
    Close #f
End Function

Sub SaveAllPlayersOnline()
    Dim I As Long

    For I = 1 To HighIndex
        If IsPlaying(I) Then
            Call SavePlayer(I)
        End If
    Next I
End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim I As Long

    Call CheckClasses

    FileName = App.Path & "\data\classes.ini"

    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))

    ReDim Class(0 To Max_Classes) As ClassRec

    Call ClearClasses

    For I = 0 To Max_Classes
        Class(I).Name = GetVar(FileName, "CLASS" & I, "Name")
        Class(I).Sprite = GetVar(FileName, "CLASS" & I, "Sprite")
        Class(I).FSprite = GetVar(FileName, "CLASS" & I, "FSprite")
        Class(I).STR = Val(GetVar(FileName, "CLASS" & I, "STR"))
        Class(I).DEF = Val(GetVar(FileName, "CLASS" & I, "DEF"))
        Class(I).SPEED = Val(GetVar(FileName, "CLASS" & I, "SPEED"))
        Class(I).MAGI = Val(GetVar(FileName, "CLASS" & I, "MAGI"))

        DoEvents
    Next I
End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim I As Long

    FileName = App.Path & "\data\classes.ini"

    For I = 0 To Max_Classes
        Call PutVar(FileName, "CLASS" & I, "Name", Trim$(Class(I).Name))
        Call PutVar(FileName, "CLASS" & I, "Sprite", STR$(Class(I).Sprite))
        Call PutVar(FileName, "CLASS" & I, "FSprite", STR$(Class(I).FSprite))
        Call PutVar(FileName, "CLASS" & I, "STR", STR$(Class(I).STR))
        Call PutVar(FileName, "CLASS" & I, "DEF", STR$(Class(I).DEF))
        Call PutVar(FileName, "CLASS" & I, "SPEED", STR$(Class(I).SPEED))
        Call PutVar(FileName, "CLASS" & I, "MAGI", STR$(Class(I).MAGI))
    Next I
End Sub

Sub CheckClasses()
    If Not FileExist("data\classes.ini") Then
        Call SaveClasses
    End If
End Sub

Sub SaveItems()
    Dim I As Long

    Call SetStatus("Saving items... ")

    For I = 1 To MAX_ITEMS

        If Not FileExist("data\items\item" & I & ".itm") Then
            Call SetStatus("Saving items... ")

            DoEvents
            Call SaveItem(I)
        End If

    Next

End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\data\items\item" & ItemNum & ".itm"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckItems

    For I = 1 To MAX_ITEMS
        Call SetStatus("Loading items... ")
        FileName = App.Path & "\data\items\item" & I & ".itm"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(I)
        Close #f

        DoEvents
    Next

End Sub

Sub CheckItems()
    Call SaveItems
End Sub

' Begin Guilds

Sub SaveGuilds()
    Dim I As Long

    Call SetStatus("Saving Guilds... ")

    For I = 1 To MAX_GUILDS

        If Not FileExist("data\guilds\guild" & I & ".gld") Then
            Call SetStatus("Saving Guilds... ")

            DoEvents
            Call SaveGuild(I)
        End If

    Next

End Sub

Sub SaveGuild(ByVal GuildNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\data\guilds\guild" & GuildNum & ".gld"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Guild(GuildNum)
    Close #f
End Sub

Sub LoadGuilds()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckGuilds

    For I = 1 To MAX_GUILDS
        Call SetStatus("Loading Guilds... ")
        FileName = App.Path & "\data\guilds\guild" & I & ".gld"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Guild(I)
        Close #f

        DoEvents
    Next

End Sub

Sub CheckGuilds()
    Call SaveGuilds
End Sub

' End Guilds

'Begin Quests

Sub SaveQuests()
    Dim I As Long

    Call SetStatus("Saving Quests... ")

    For I = 1 To MAX_QUESTS

        If Not FileExist("data\quests\quest" & I & ".qst") Then
            Call SetStatus("Saving Quests... ")

            DoEvents
            Call SaveQuest(I)
        End If

    Next

End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\data\quests\quest" & QuestNum & ".qst"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Quest(QuestNum)
    Close #f
End Sub

Sub LoadQuests()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckQuests

    For I = 1 To MAX_QUESTS
        Call SetStatus("Loading Quests... ")
        FileName = App.Path & "\data\quests\quest" & I & ".qst"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Quest(I)
        Close #f

        DoEvents
    Next

End Sub

Sub CheckQuests()
    Call SaveQuests
End Sub

'End Quests

Sub SaveShops()
    Dim I As Long

    Call SetStatus("Saving Shops... ")

    For I = 1 To MAX_SHOPS

        If Not FileExist("data\Shops\Shop" & I & ".shp") Then
            Call SetStatus("Saving Shops... ")

            DoEvents
            Call SaveShop(I)
        End If

    Next

End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\data\Shops\Shop" & ShopNum & ".shp"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckShops

    For I = 1 To MAX_SHOPS
        Call SetStatus("Loading Shops... ")
        FileName = App.Path & "\data\Shops\Shop" & I & ".shp"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Shop(I)
        Close #f

        DoEvents
    Next

End Sub

Sub CheckShops()
    Call SaveShops
End Sub

Sub SaveSign(ByVal SignNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\data\Signs\Sign" & SignNum & ".sign"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Sign(SignNum)
    Close #f
End Sub

Sub SaveSigns()
    Dim I As Long

    Call SetStatus("Saving Signs... ")

    For I = 1 To MAX_SIGNS

        If Not FileExist("data\Signs\Sign" & I & ".sign") Then
            Call SetStatus("Saving Signs... ")

            DoEvents
            Call SaveSign(I)
        End If

    Next

End Sub

Sub LoadSigns()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckSigns

    For I = 1 To MAX_SIGNS
        Call SetStatus("Loading signs... ")
        FileName = App.Path & "\data\Signs\Sign" & I & ".sign"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Sign(I)
        Close #f
        If I Mod 20 Then DoEvents
    Next
End Sub

Sub CheckSigns()
    Call SaveSigns
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\data\Spells\Spell" & SpellNum & ".spl"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim I As Long

    Call SetStatus("Saving Spells... ")

    For I = 1 To MAX_SPELLS

        If Not FileExist("data\Spells\Spell" & I & ".spl") Then
            Call SetStatus("Saving Spells... ")

            DoEvents
            Call SaveSpell(I)
        End If

    Next

End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckSpells

    For I = 1 To MAX_SPELLS
        Call SetStatus("Loading Spells... ")
        FileName = App.Path & "\data\Spells\Spell" & I & ".spl"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Spell(I)
        Close #f

        DoEvents
    Next

End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\data\Npcs\Npc" & NpcNum & ".npc"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub SaveNpcs()
    Dim I As Long

    Call SetStatus("Saving Npcs... ")

    For I = 1 To MAX_NPCS

        If Not FileExist("data\Npcs\Npc" & I & ".npc") Then
            Call SetStatus("Saving Npcs... ")

            DoEvents
            Call SaveNpc(I)
        End If

    Next

End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckNpcs

    For I = 1 To MAX_NPCS
        Call SetStatus("Loading Npcs... ")
        FileName = App.Path & "\data\Npcs\Npc" & I & ".npc"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Npc(I)
        Close #f

        DoEvents
    Next

End Sub

Sub CheckNpcs()
    Call SaveNpcs
End Sub

Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SaveMaps()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    For I = 1 To MAX_MAPS_SET
        Call SaveMap(I)
    Next I
End Sub

Sub LoadMaps()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckMaps

    For I = 1 To MAX_MAPS_SET
        FileName = App.Path & "\maps\map" & I & ".dat"

        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Map(I)
        Close #f

        DoEvents
    Next I
End Sub

Sub ConvertOldMapsToNew()
    Dim FileName As String
    Dim I As Long
    Dim f As Long
    Dim X As Long, y As Long
    Dim OldMap As OldMapRec
    Dim NewMap As MapRec

    For I = 1 To MAX_MAPS_SET
        FileName = App.Path & "\maps\map" & I & ".dat"

        ' Get the old file
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , OldMap
        Close #f

        ' Delete the old file
        Call Kill(FileName)

        ' Convert
        NewMap.Name = OldMap.Name
        NewMap.Revision = OldMap.Revision + 1
        NewMap.Moral = OldMap.Moral
        NewMap.Up = OldMap.Up
        NewMap.Down = OldMap.Down
        NewMap.Left = OldMap.Left
        NewMap.Right = OldMap.Right
        NewMap.Music = OldMap.Music
        NewMap.BootMap = OldMap.BootMap
        NewMap.BootX = OldMap.BootX
        NewMap.BootY = OldMap.BootY
        NewMap.Shop = OldMap.Shop
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                NewMap.Tile(X, y).Ground = OldMap.Tile(X, y).Ground
                NewMap.Tile(X, y).Mask = OldMap.Tile(X, y).Mask
                NewMap.Tile(X, y).Anim = OldMap.Tile(X, y).Anim
                NewMap.Tile(X, y).Fringe = OldMap.Tile(X, y).Fringe
                NewMap.Tile(X, y).Type = OldMap.Tile(X, y).Type
                NewMap.Tile(X, y).Data1 = OldMap.Tile(X, y).Data1
                NewMap.Tile(X, y).Data2 = OldMap.Tile(X, y).Data2
                NewMap.Tile(X, y).Data3 = OldMap.Tile(X, y).Data3
            Next X
        Next y

        For X = 1 To MAX_MAP_NPCS
            NewMap.Npc(X) = OldMap.Npc(X)
        Next X

        ' Set new values to 0 or null
        NewMap.Indoors = NO

        ' Save the new map
        f = FreeFile
        Open FileName For Binary As #f
        Put #f, , NewMap
        Close #f
    Next I
End Sub

Sub CheckMaps()
    Dim FileName As String
    Dim X As Long
    Dim y As Long
    Dim I As Long
    Dim N As Long

    Call ClearMaps

    For I = 1 To MAX_MAPS_SET
        FileName = "maps\map" & I & ".dat"

        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            Call SaveMap(I)
        End If
    Next I
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\logs\" & FN

        If Not FileExist("logs\" & FN) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If

        f = FreeFile
        Open FileName For Append As #f
        Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, Optional BannedByIndex As String)
    Dim FileName As String, IP As String, bnam As String
    Dim f As Long, I As Long
    Dim BNum As Integer, b As Integer
    Dim MAX_BANS As Long

    FileName = App.Path & "\banlist.ini"

    IP = GetPlayerIP(BanPlayerIndex)

    b = 1

    For I = 0 To MAX_BANS + 1
        If Ban(I).BannedIP = vbNullString Then
            BNum = I
            Exit For
        End If

        If I = MAX_BANS + 1 Then
            BNum = MAX_BANS + 1
            Exit For
        End If

    Next I

    ' Add there data to a ban slot
    Ban(BNum).BannedIP = IP
    Ban(BNum).BannedChar = GetPlayerName(BanPlayerIndex)
    Ban(BNum).BannedBy = Trim$(BannedByIndex)
    Ban(BNum).BannedHD = GetPlayerHD(BanPlayerIndex)
    Call SaveBan(BNum)
    Call PutVar(FileName, "Total", "Total", GetVar(FileName, "Total", "Total") + 1)
    MAX_BANS = MAX_BANS + 1

' Alert People
End Sub

Sub UnBanIndex(ByVal BannedPlayerName As String, ByVal DeBannedByIndex As Long)
    Dim FileName As String, IP As String, bnam As String
    Dim f As Long, I As Long
    Dim b As Integer
    Dim MAX_BANS As Integer

    FileName = App.Path & "\banlist.ini"

    For I = 0 To MAX_BANS + 1
        If LCase$(GetVar(FileName, "Ban" & I, "BannedChar")) = LCase$(BannedPlayerName) Then
            ' Delete there data to a ban slot
            Ban(I).BannedIP = vbNullString
            Ban(I).BannedChar = vbNullString
            Ban(I).BannedBy = vbNullString
            Ban(I).BannedHD = vbNullString
            Call SaveBan(I)

            ' Alert People
            Call GlobalMsg(BannedPlayerName & " has been unbanned from " & GAME_NAME & " by " & GetPlayerName(DeBannedByIndex) & "!", White)
            Call AddLog(GetPlayerName(DeBannedByIndex) & " has unbanned " & BannedPlayerName & ".", ADMIN_LOG)
            Exit For
        End If

        If I = MAX_BANS + 1 Then
            Call PlayerMsg(DeBannedByIndex, "Player is not banned!", White)
        End If
    Next I

End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long, f2 As Long
    Dim s As String

    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")

    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s
        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If
    Loop

    Close #f1
    Close #f2

    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

Sub LoadBans()
    Dim FileName As String
    Dim I As Integer
    Dim BNum As Integer
    FileName = App.Path & "\data\banlist.ini"
    MAX_BANS = GetVar(FileName, "Total", "Total")

    ReDim Ban(0 To MAX_BANS) As BanRec

    For I = 0 To MAX_BANS

        If GetVar(FileName, "Ban" & I, "BannedIP") = vbNullString Then
            Ban(I).BannedIP = GetVar(FileName, "Ban" & I, "BannedIP")
            Ban(I).BannedChar = GetVar(FileName, "Ban" & I, "BannedChar")
            Ban(I).BannedBy = GetVar(FileName, "Ban" & I, "BannedBy")
            Ban(I).BannedHD = GetVar(FileName, "Ban" & I, "BannedHD")
        End If

    Next I

End Sub

Sub SaveBan(ByVal BanNum As Integer)
    Dim FileName As String
    FileName = App.Path & "\banlist.ini"

    Call PutVar(FileName, "Ban" & BanNum, "BannedIP", Ban(BanNum).BannedIP)
    Call PutVar(FileName, "Ban" & BanNum, "BannedChar", Ban(BanNum).BannedChar)
    Call PutVar(FileName, "Ban" & BanNum, "BannedBy", Ban(BanNum).BannedBy)
    Call PutVar(FileName, "Ban" & BanNum, "BannedHD", Ban(BanNum).BannedHD)

End Sub
