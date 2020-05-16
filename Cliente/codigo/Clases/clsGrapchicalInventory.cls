VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MouseDownSelectedItem      As Long
Private InvSelectedItem            As Long             'Currently selected item

Private ownTilePixelWidth          As Integer
Private ownTilePixelHeight         As Integer

Private startdX                    As Integer
Private startdY                    As Integer

Private ShowValue                  As Boolean
Private ValueFormated              As Boolean
Private ImgContainer               As Boolean

Private InventoryRect              As D3DRECT
Private w                          As Long
Private H                          As Long
Private Inventory()                As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Public Sub SelectGold()

    Dim prevSelectedItem As Long
    
    prevSelectedItem = InvSelectedItem
    
    InvSelectedItem = FLAGORO
    
    If prevSelectedItem <> FLAGORO Then
        RenderInv
    End If

End Sub

Public Sub DeselectItem()

    If InvSelectedItem = 0 Then Exit Sub
    
    InvSelectedItem = 0
    
    RenderInv

End Sub

Public Sub ChangeSlotItemAmount(ByVal Slot As Byte, ByVal Amount As Long)
    
    Inventory(Slot).Amount = Amount

    RenderInv

End Sub

Public Property Get SelectedItem() As Long

    SelectedItem = InvSelectedItem

End Property

Public Property Get MaxHit(ByVal Slot As Byte) As Integer

    MaxHit = Inventory(Slot).MaxHit

End Property

Public Property Get MinHit(ByVal Slot As Byte) As Integer

    MinHit = Inventory(Slot).MinHit

End Property

Public Property Get MaxDef(ByVal Slot As Byte) As Integer

    MaxDef = Inventory(Slot).MaxDef

End Property

Public Property Get MinDef(ByVal Slot As Byte) As Integer

    MinDef = Inventory(Slot).MinDef

End Property

Public Property Get GrhIndex(ByVal Slot As Byte) As Long

    GrhIndex = Inventory(Slot).GrhIndex

End Property

Public Property Get Valor(ByVal Slot As Byte) As Single

    Valor = Inventory(Slot).Valor

End Property

Public Property Get Amount(ByVal Slot As Byte) As Long

    If Slot = FLAGORO Then
        Amount = UserGLD
    ElseIf Slot >= LBound(Inventory) And Slot <= UBound(Inventory) Then
        Amount = Inventory(Slot).Amount
    End If

End Property

Public Property Get ObjIndex(ByVal Slot As Byte) As Integer

    ObjIndex = Inventory(Slot).ObjIndex

End Property

Public Property Get ObjType(ByVal Slot As Byte) As Integer

    ObjType = Inventory(Slot).ObjType

End Property

Public Property Get ItemName(ByVal Slot As Byte) As String

    ItemName = Inventory(Slot).Name

End Property

Public Property Get Equipped(ByVal Slot As Byte) As Boolean

    Equipped = Inventory(Slot).Equipped

End Property

Public Property Get MaxObjs() As Byte

    MaxObjs = UBound(Inventory)

End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)

    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory

End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)

    ShowValue = bShow
    
    Dim Slot As Long

    For Slot = 1 To Me.MaxObjs
        RenderInv
    Next Slot

End Sub

Public Sub ClearAllSlots()

    Dim Slot As Long

    For Slot = 1 To Me.MaxObjs
        Call SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
    Next Slot

End Sub

Public Sub SetItem(ByVal Slot As Byte, _
    ByVal eOBJIndex As Integer, _
    ByVal eAmount As Long, _
    ByVal eEquipped As Byte, _
    ByVal eGrhIndex As Integer, _
    ByVal eObjType As Integer, _
    ByVal eMaxHit As Integer, _
    ByVal eMinHit As Integer, _
    ByVal eMaxDef As Integer, _
    ByVal eMinDef As Integer, _
    ByVal eValor As Single, _
    ByVal eName As String)

    If Slot < 1 Or Slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(Slot)
        .Amount = eAmount
        .MaxDef = eMaxDef
        .MinDef = eMinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .ObjIndex = eOBJIndex
        .ObjType = eObjType
        .Valor = eValor

    End With
    
    Exit Sub

End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long

    Dim TempItem As Long
    Dim Temp_X   As Long
    Dim Temp_Y   As Long
    
    Temp_X = X \ ownTilePixelWidth
    Temp_Y = Y \ ownTilePixelHeight
    
    TempItem = Temp_X + Temp_Y * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem

    End If

End Function

Public Sub RenderInv()

    Dim LooPC As Long

    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    DirectDevice.BeginScene
    
    For LooPC = 1 To UBound(Inventory)
        startdX = ((LooPC - 1) Mod w) * ownTilePixelHeight
        startdY = ((LooPC - 1) \ H) * ownTilePixelHeight

        If Inventory(LooPC).GrhIndex Then
       
            If InvSelectedItem = LooPC Then
                If ImgContainer Then
                    Directx_Render_Texture CLng(2), startdX, startdY, 32, 32, 64, 0, White

                End If

            End If
            
            Directx_Render_Texture CLng(GrhData(Inventory(LooPC).GrhIndex).FileNum), startdX, startdY, 32, 32, GrhData(Inventory( _
                LooPC).GrhIndex).sX, GrhData(Inventory(LooPC).GrhIndex).sY, White

            If Inventory(LooPC).Equipped Then
                Text_Draw startdX + 19, startdY + 21, "+", White

            End If
           
            If ShowValue Then

                Dim sValue As String

               'sValue = Replace$(Inventory(LoopC).Amount, "0000", "0K")
               
               sValue = Inventory(LooPC).Amount
            
                Text_Draw startdX, startdY, sValue, Yellow

            End If
    
        End If
        
    Next LooPC
            
    Call Directx_EndScene(InventoryRect, InventoryWindow.hwnd)

End Sub

Public Sub Initialize(ByRef InvPic As PictureBox, _
    ByVal MaxObjs As Byte, _
    Optional ByVal FontSize As Integer = 7, _
    Optional ByVal TileWidth As Integer = 32, _
    Optional ByVal TileHeight As Integer = 32, _
    Optional ByVal startX As Integer = 0, _
    Optional ByVal startY As Integer = 0, _
    Optional ByVal bImgContainer As Boolean = True, _
    Optional ByVal bShowText As Boolean = True, _
    Optional ByVal FormatValue As Boolean = True)

    Set InventoryWindow = InvPic
    ValueFormated = FormatValue
    
    InventoryWindow.AutoRedraw = True

    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    
    startdX = startX
    startdY = startY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText

    With InventoryRect
        .Y2 = InventoryWindow.ScaleHeight
        .X2 = InventoryWindow.ScaleWidth

    End With
    
    w = InventoryWindow.ScaleWidth / ownTilePixelWidth
    H = InventoryWindow.ScaleWidth / ownTilePixelHeight
       
    If ImgContainer Then InvSelectedItem = ClickItem(1, 1)

End Sub
    
Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim TempItem As Integer
    
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub

    TempItem = ClickItem(X, Y)
    
    If TempItem <= UBound(Inventory) And TempItem > 0 Then
        Dim tStr As String

        With Inventory(TempItem)
            tStr = .Name & " ATK: " & .MinHit & "/" & .MaxHit & " DEF: " & .MinDef & "/" & .MaxDef

        End With

        InventoryWindow.ToolTipText = tStr

    End If
    
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
                
    If (Button = vbRightButton) Then

        Dim TempItem As Long

        TempItem = ClickItem(X, Y)

        If (Inventory(TempItem).GrhIndex > 0) Then
            MouseDownSelectedItem = TempItem

            Exit Sub

        End If

    End If
   
    MouseDownSelectedItem = 0
        
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim prevSelItem As Long

    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
    
    prevSelItem = InvSelectedItem

    InvSelectedItem = ClickItem(X, Y)

    If Button = vbRightButton Then

        If (frmMain.UsandoDrag) And (InvSelectedItem <= UBound(Inventory)) And (InvSelectedItem >= 0) And (Not frmMain.UsabaDrag) Then
                
            Call SendData("DRA" & InvSelectedItem & "," & MouseDownSelectedItem)
            frmMain.UsandoDrag = False
        Else
            frmMain.UsandoDrag = False

        End If

    Else
    
        If prevSelItem <> InvSelectedItem Then
            If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call RenderInv
            If InvSelectedItem Then Call RenderInv

        End If

    End If

End Sub
