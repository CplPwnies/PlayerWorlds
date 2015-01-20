Attribute VB_Name = "modDirectX"
Option Explicit

Public Sub InitDirectX()
Attribute InitDirectX.VB_UserMemId = 1610612736
' ****************************************************************
' * WHEN    WHO    WHAT
' * ----    ---    ----
' * 07/12/2005  Shannara   Optimized function.
' ****************************************************************

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate(vbNullString)

    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMainGame.hwnd, DDSCL_NORMAL)

    ' Init type and get the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    End With

    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)

    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMainGame.picScreen.hwnd

    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip

    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Public Sub InitSurfaces()
Attribute InitSurfaces.VB_UserMemId = 1610612737
' ****************************************************************
' * WHEN    WHO    WHAT
' * ----    ---    ----
' * 07/12/2005  Shannara   Optimized function, added gfx constants.
' ****************************************************************

    Dim Key As DDCOLORKEY
    Dim FileName As String

    ' Set path prefix
    FileName = App.Path & GFX_PATH

    ' Check for files existing
    If FileExist(FileName & "sprites" & GFX_EXT, True) = False Or FileExist(FileName & "tiles" & GFX_EXT, True) = False Or FileExist(FileName & "items" & GFX_EXT, True) = False Or FileExist(FileName & "spells" & GFX_EXT, True) = False Then
        Call MsgBox("You dont have the graphics files in the " & FileName & " directory!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If

    ' Set the key for masks
    With Key
        .low = 0
        .high = 0
    End With

    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lWidth = (MAX_MAPX + 1) * PIC_X
        .lHeight = (MAX_MAPY + 1) * PIC_Y
    End With
    
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Set DD_LowerBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Set DD_MiddleBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Set DD_UpperBuffer = DD.CreateSurface(DDSD_BackBuffer)
    DD_BackBuffer.SetColorKey DDCKEY_SRCBLT, Key
    DD_LowerBuffer.SetColorKey DDCKEY_SRCBLT, Key
    DD_MiddleBuffer.SetColorKey DDCKEY_SRCBLT, Key
    DD_UpperBuffer.SetColorKey DDCKEY_SRCBLT, Key

    ' Init sprite ddsd type and load the bitmap
    With DDSD_Sprite
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(FileName & "sprites" & GFX_EXT, DDSD_Sprite)
    ' SetMaskColorFromPixel DD_SpriteSurf, 0, 0
    DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, Key

    ' Init tiles ddsd type and load the bitmap
    With DDSD_Tile
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    Set DD_TileSurf = DD.CreateSurfaceFromFile(FileName & "tiles" & GFX_EXT, DDSD_Tile)
    ' SetMaskColorFromPixel DD_TileSurf, 0, 0
    DD_TileSurf.SetColorKey DDCKEY_SRCBLT, Key

    ' Init items ddsd type and load the bitmap
    With DDSD_Item
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(FileName & "items" & GFX_EXT, DDSD_Item)
    ' SetMaskColorFromPixel DD_ItemSurf, 0, 0
    DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, Key

    ' Init spells ddsd type and load the bitmap
    With DDSD_Spell
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    Set DD_SpellSurf = DD.CreateSurfaceFromFile(FileName & "spells" & GFX_EXT, DDSD_Spell)
    ' SetMaskColorFromPixel DD_SpellSurf, 0, 0
    DD_SpellSurf.SetColorKey DDCKEY_SRCBLT, Key

End Sub

Sub DestroyDirectX()
Attribute DestroyDirectX.VB_UserMemId = 1610612738
    Set DX = Nothing
    Set DD = Nothing
    Set DD_Clip = Nothing

    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_SpellSurf = Nothing
    Set DD_BackBuffer = Nothing

    Set DD_LowerBuffer = Nothing
    Set DD_MiddleBuffer = Nothing
    Set DD_UpperBuffer = Nothing
End Sub

Function NeedToRestoreSurfaces() As Boolean
Attribute NeedToRestoreSurfaces.VB_UserMemId = 1610612739
    Dim TestCoopRes As Long

    TestCoopRes = DD.TestCooperativeLevel

    If (TestCoopRes = DD_OK) Then
        NeedToRestoreSurfaces = False
    Else
        NeedToRestoreSurfaces = True
    End If
End Function

Public Sub CheckSurfaces()
Attribute CheckSurfaces.VB_UserMemId = 1610612740
    On Error GoTo ErrorHandle

    ' Check if we need to restore surfaces
    If NeedToRestoreSurfaces Then
        DD.RestoreAllSurfaces
        Call InitSurfaces
    End If

    Exit Sub

ErrorHandle:
    Call GameDestroy

End Sub

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Attribute SetMaskColorFromPixel.VB_UserMemId = 1610612741
    Dim TmpR As RECT
    Dim TmpDDSD As DDSURFACEDESC2
    Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = X
        .top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .low = TheSurface.GetLockedPixel(X, Y)
        .high = .low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

    TheSurface.Unlock TmpR
End Sub


Public Sub BltMap()
' ****************************************************************
' * WHEN    WHO    WHAT
' * ----    ---    ----
' * 07/12/2005  Shannara   Optimized function.
' * 04/10/2007  Robin    Added awesomeness :)
' ****************************************************************

    Dim Ground As Long
    Dim Anim1 As Long
    Dim Anim2 As Long
    Dim Mask2 As Long
    Dim M2Anim As Long
    Dim Fringe As Long
    Dim FAnim As Long
    Dim Fringe2 As Long
    Dim F2Anim As Long
    Dim X As Long, Y As Long

    rec.top = 0
    rec.Bottom = (MAX_MAPY + 1) * 32
    rec.Left = 0
    rec.Right = (MAX_MAPX + 1) * 32
    
    DD_LowerBuffer.BltColorFill rec, RGB(0, 0, 0)
    DD_UpperBuffer.BltColorFill rec, RGB(0, 0, 0)

    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY

            With Map.Tile(X, Y)
                Ground = Map.Tile(X, Y).Ground
                Anim1 = Map.Tile(X, Y).Mask
                Anim2 = Map.Tile(X, Y).Anim
                Mask2 = Map.Tile(X, Y).Mask2
                M2Anim = Map.Tile(X, Y).M2Anim
                Fringe = Map.Tile(X, Y).Fringe
                FAnim = Map.Tile(X, Y).FAnim
                Fringe2 = Map.Tile(X, Y).Fringe2
                F2Anim = Map.Tile(X, Y).F2Anim
            End With

            With rec
                .top = Int(Ground / 7) * PIC_Y
                .Bottom = .top + PIC_Y
                .Left = (Ground - Int(Ground / 7) * 7) * PIC_X
                .Right = .Left + PIC_X
            End With
            Call DD_LowerBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT)

            If (MapAnim = 0) Or (Anim2 <= 0) Then
                ' Is there an animation tile to plot?
                If Anim1 > 0 And TempTile(X, Y).DoorOpen = NO Then
                    rec.top = Int(Anim1 / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (Anim1 - Int(Anim1 / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_LowerBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            Else
                ' Is there a second animation tile to plot?
                If Anim2 > 0 Then
                    rec.top = Int(Anim2 / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (Anim2 - Int(Anim2 / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_LowerBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If (MapAnim = 0) Or (M2Anim <= 0) Then
                ' Is there an animation tile to plot?
                If Mask2 > 0 Then
                    rec.top = Int(Mask2 / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (Mask2 - Int(Mask2 / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_LowerBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            Else
                ' Is there a second animation tile to plot?
                If M2Anim > 0 Then
                    rec.top = Int(M2Anim / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (M2Anim - Int(M2Anim / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_LowerBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If (MapAnim = 0) Or (FAnim <= 0) Then
                ' Is there an animation tile to plot?

                If Fringe > 0 Then
                    rec.top = Int(Fringe / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (Fringe - Int(Fringe / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_UpperBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

            Else

                If FAnim > 0 Then
                    rec.top = Int(FAnim / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (FAnim - Int(FAnim / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_UpperBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

            End If

            If (MapAnim = 0) Or (F2Anim <= 0) Then
            ' Is there an animation tile to plot?

                If Fringe2 > 0 Then
                    rec.top = Int(Fringe2 / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (Fringe2 - Int(Fringe2 / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_UpperBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

            Else

                If F2Anim > 0 Then
                    rec.top = Int(F2Anim / 7) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = (F2Anim - Int(F2Anim / 7) * 7) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DD_UpperBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

            End If

        Next
    Next
End Sub


Public Sub BltItem(ByVal ItemNum As Long)
' ****************************************************************
' * WHEN    WHO    WHAT
' * ----    ---    ----
' * 07/12/2005  Shannara   Optimized function.
' ****************************************************************

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = MapItem(ItemNum).Y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = MapItem(ItemNum).X * PIC_X
        .Right = .Left + PIC_X
    End With

    With rec
        .top = Item(MapItem(ItemNum).Num).Pic * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    Call DD_MiddleBuffer.BltFast(MapItem(ItemNum).X * PIC_X, MapItem(ItemNum).Y * PIC_Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltPlayer(ByVal index As Long)
    ' ****************************************************************
    ' * WHEN    WHO    WHAT
    ' * ----    ---    ----
    ' * 07/12/2005  Shannara   Optimized function.
    ' ****************************************************************
    Dim X As Long, Y As Long

    ' Check for player(index).animation
    If Player(index).Attacking = 0 Then
        Select Case GetPlayerDir(index)
            Case DIR_UP
                Player(index).Anim = 0
                If (Player(index).YOffset < PIC_Y / 3) Then
                    Player(index).Anim = 1
                ElseIf (Player(index).YOffset > PIC_Y / 3) And ((Player(index).YOffset > PIC_Y / 3 * 2)) Then
                    Player(index).Anim = 2
                End If
            Case DIR_DOWN
                Player(index).Anim = 1
                If (Player(index).YOffset < PIC_X / 4 * -1) Then Player(index).Anim = 0
                If (Player(index).YOffset < PIC_X / 2 * -1) Then Player(index).Anim = 2
            Case DIR_LEFT
                Player(index).Anim = 0
                If (Player(index).XOffset < PIC_Y / 3) Then
                    Player(index).Anim = 1
                ElseIf (Player(index).XOffset > PIC_Y / 3) And ((Player(index).XOffset > PIC_Y / 3 * 2)) Then
                    Player(index).Anim = 2
                End If
            Case DIR_RIGHT
                Player(index).Anim = 0
                If (Player(index).XOffset < PIC_Y / 4 * -1) Then Player(index).Anim = 1
                If (Player(index).XOffset < PIC_Y / 2 * -1) Then Player(index).Anim = 2
        End Select
    ' Dim obj As pwMovement.clsWalk
    ' Set obj = CreateObject("pwMovement.clsWalk")
    ' Player(index).Anim = obj.GetPlayerAnim(GetPlayerDir(index), Player(index).YOffset, Player(index).XOffset)
    ' Set obj = Nothing
    Else
        If Player(index).AttackTimer + 500 > GetTickCount Then
            Player(index).Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    With Player(index)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    With rec
        .top = (GetPlayerSprite(index) * GameData.PlayerY)
        .Bottom = .top + PIC_Y
        .Left = (GetPlayerDir(index) * 3 + Player(index).Anim) * GameData.PlayerX
        .Right = .Left + GameData.PlayerX
    End With

    If GameData.PlayerX > 32 Then
        X = (GetPlayerX(index) * PIC_X) + (Player(index).XOffset) - (GameData.PlayerX / 4)
    Else
        X = (GetPlayerX(index) * PIC_X) + (Player(index).XOffset)
    End If
    Y = (GetPlayerY(index) * PIC_Y) + (Player(index).YOffset)

    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .top = .top + (Y * -1)
        End With
    End If

    Call DD_MiddleBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltPlayerTop(ByVal index As Long)
    Dim X As Long, Y As Long

    With rec
        .top = (GetPlayerSprite(index) * GameData.PlayerY) - GameData.PlayerY
        .Bottom = .top + (GameData.PlayerY - 32)
        .Left = (GetPlayerDir(index) * 3 + Player(index).Anim) * GameData.PlayerX
        .Right = .Left + GameData.PlayerX
    End With

    If GameData.PlayerX > 32 Then
        X = (GetPlayerX(index) * PIC_X) + (Player(index).XOffset) - (GameData.PlayerX / 4)
    Else
        X = (GetPlayerX(index) * PIC_X) + (Player(index).XOffset)
    End If
    Y = (GetPlayerY(index) * PIC_Y) + (Player(index).YOffset)

    Y = Y - (GameData.PlayerY - 32)

    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .top = .top + (Y * -1)
        End With
    End If

    Call DD_MiddleBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub SpellEditorBltAnim(ByVal Frame As Byte)
    Call BitBlt(frmSpellEditor.picAnim.hDC, 0, 0, PIC_X, PIC_Y, frmSpellEditor.picSpells.hDC, Frame * PIC_X, frmSpellEditor.scrlAnim.Value * PIC_Y, SRCCOPY)
End Sub

Sub BltSpell(ByVal VicX As Long, ByVal VicY As Long, ByVal SpellAnim As Byte)
    If SpellVar > 13 Then
        Exit Sub
    End If
    ' Change Spell Animation Every 250 miliseconds
    If GetTickCount > SpellAnimTimer + 75 Then
        If SpellVar > 13 Then
            SpellVar = 0
            Exit Sub
        Else
            SpellVar = SpellVar + 1
        End If
        SpellAnimTimer = GetTickCount
    End If

    ' 32x32 Spells
    rec.top = SpellAnim * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = SpellVar * PIC_X
    rec.Right = rec.Left + PIC_X

    ' 32x32 spells
    Call DD_MiddleBuffer.BltFast(VicX * PIC_X, VicY * PIC_Y, DD_SpellSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
' ****************************************************************
' * WHEN    WHO    WHAT
' * ----    ---    ----
' * 07/12/2005  Shannara   Optimized function.
' ****************************************************************

    Dim Anim As Byte
    Dim X As Long, Y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .top + PIC_Y
        .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
        .Right = .Left + PIC_X
    End With

    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    With rec
        .top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With

    With MapNpc(MapNpcNum)
        X = .X * PIC_X + .XOffset
        Y = .Y * PIC_Y + .YOffset - 4
    End With

    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .top = .top + (Y * -1)
        End With
    End If

    Call DD_MiddleBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub



