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
Private h                          As Long
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

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal Amount As Long)
    
    Inventory(slot).Amount = Amount

    RenderInv

End Sub

Public Property Get SelectedItem() As Long

    SelectedItem = InvSelectedItem

End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer

    MaxHit = Inventory(slot).MaxHit

End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer

    MinHit = Inventory(slot).MinHit

End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer

    MaxDef = Inventory(slot).MaxDef

End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer

    MinDef = Inventory(slot).MinDef

End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Long

    GrhIndex = Inventory(slot).GrhIndex

End Property

Public Property Get Valor(ByVal slot As Byte) As Single

    Valor = Inventory(slot).Valor

End Property

Public Property Get Amount(ByVal slot As Byte) As Long

    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        Amount = Inventory(slot).Amount
    End If

End Property

Public Property Get ObjIndex(ByVal slot As Byte) As Integer

    ObjIndex = Inventory(slot).ObjIndex

End Property

Public Property Get ObjType(ByVal slot As Byte) As Integer

    ObjType = Inventory(slot).ObjType

End Property

Public Property Get ItemName(ByVal slot As Byte) As String

    ItemName = Inventory(slot).Name

End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean

    Equipped = Inventory(slot).Equipped

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
    
    Dim slot As Long

    For slot = 1 To Me.MaxObjs
        RenderInv
    Next slot

End Sub

Public Sub ClearAllSlots()

    Dim slot As Long

    For slot = 1 To Me.MaxObjs
        Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
    Next slot

End Sub

Public Sub SetItem(ByVal slot As Byte, _
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

    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
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

Private Function ClickItem(ByVal x As Long, ByVal y As Long) As Long

    Dim TempItem As Long
    Dim Temp_X   As Long
    Dim Temp_Y   As Long
    
    Temp_X = x \ ownTilePixelWidth
    Temp_Y = y \ ownTilePixelHeight
    
    TempItem = Temp_X + Temp_Y * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem

    End If

End Function

Public Sub RenderInv()

    Dim LoopC As Long

    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    DirectDevice.BeginScene
    
    For LoopC = 1 To UBound(Inventory)
        startdX = ((LoopC - 1) Mod w) * ownTilePixelHeight
        startdY = ((LoopC - 1) \ h) * ownTilePixelHeight

        If Inventory(LoopC).GrhIndex Then
       
            If InvSelectedItem = LoopC Then
                If ImgContainer Then
                    Directx_Render_Texture CLng(2), startdX, startdY, 32, 32, 64, 0, White

                End If

            End If
            
            Directx_Render_Texture CLng(GrhData(Inventory(LoopC).GrhIndex).FileNum), startdX, startdY, 32, 32, GrhData(Inventory( _
                LoopC).GrhIndex).sX, GrhData(Inventory(LoopC).GrhIndex).sY, White

            If Inventory(LoopC).Equipped Then
                Text_Draw startdX + 19, startdY + 21, "+", White

            End If
           
            If ShowValue Then

                Dim sValue As String

               'sValue = Replace$(Inventory(LoopC).Amount, "0000", "0K")
               
               sValue = Inventory(LoopC).Amount
            
                Text_Draw startdX, startdY, sValue, Yellow

            End If
    
        End If
        
    Next LoopC
            
    Call Directx_EndScene(InventoryRect, InventoryWindow.HWnd)

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
    h = InventoryWindow.ScaleWidth / ownTilePixelHeight
       
    If ImgContainer Then InvSelectedItem = ClickItem(1, 1)

End Sub
    
Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim TempItem As Integer
    
    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then Exit Sub

    TempItem = ClickItem(x, y)
    
    If TempItem <= UBound(Inventory) And TempItem > 0 Then
        Dim tStr As String

        With Inventory(TempItem)
            tStr = .Name & " ATK: " & .MinHit & "/" & .MaxHit & " DEF: " & .MinDef & "/" & .MaxDef

        End With

        InventoryWindow.ToolTipText = tStr

    End If
    
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then Exit Sub
                
    If (Button = vbRightButton) Then

        Dim TempItem As Long

        TempItem = ClickItem(x, y)

        If (Inventory(TempItem).GrhIndex > 0) Then
            MouseDownSelectedItem = TempItem

            Exit Sub

        End If

    End If
   
    MouseDownSelectedItem = 0
        
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim prevSelItem As Long

    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then Exit Sub
    
    prevSelItem = InvSelectedItem

    InvSelectedItem = ClickItem(x, y)

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
