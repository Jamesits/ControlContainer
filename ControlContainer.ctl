VERSION 5.00
Begin VB.UserControl ControlContainer 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ControlContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'事件声明:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "当用户在拥有焦点的对象上按下任意键时发生。"
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "当用户按下和释放 ANSI 键时发生。"
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "当用户在拥有焦点的对象上释放键时发生。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "当用户在拥有焦点的对象上释放鼠标发生。"
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "当控件的 Visible 属性变为 False 时发生。"
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "当 OLEDropMode 属性设置为手动、且 OLE 拖/放操作期间鼠标经过控件时发生。"
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "当移动、放大或露出图片框的任何部分时发生。"
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "当第一次显示一个窗体时或改变一个对象的大小时发生。"
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "当控件的 Visible 属性变为 True 时发生。"
'缺省属性值:
Const m_def_Alpha = 100
'属性变量:
Dim m_Alpha As Byte




'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "强制完全重画一个对象。"
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "返回/设置从 graphics 方法到一个持久性位图的输出。"
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

' "Circle" 后面的下划线是必须的，
'因为它是 VBA 中的保留字。
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Circle
Public Sub Circle_(X As Single, Y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
    UserControl.Circle (X, Y), Radius, Color, StartPos, EndPos, Aspect
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "清除窗体、图像或图片框中在运行时生成的图形和文本。"
    UserControl.Cls
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,ContainerHwnd
Public Property Get ContainerHwnd() As Long
Attribute ContainerHwnd.VB_Description = "返回句柄 (from Microsoft Windows) 到包含 UserControl 的窗口。"
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
Attribute Controls.VB_Description = "表示一个窗体上每个控件元素的集合，也包括控件数组的元素。 "
    Set Controls = UserControl.Controls
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "返回/设置下次 print 或 draw 方法的水平坐标。"
    CurrentX = UserControl.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "返回/设置下次 print 或 draw 方法的垂直坐标。"
    CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "设置以 graphics 方法或 Shape 、 Line 控件输出时的外观。"
    DrawMode = UserControl.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "决定 graphics 方法输出时的线条样式。"
    DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "返回/设置 graphics 方法输出时的线条宽度。"
    DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "返回/设置填充形状、圆环和方框所使用的颜色。"
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,FillStyle
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "返回/设置一个 shape 控件的填充样式。"
    FillStyle = UserControl.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

''注意！不要删除或修改下列被注释的行！
''MappingInfo=UserControl,UserControl,-1,FontBold
'Public Property Get FontBold() As Boolean
'    FontBold = UserControl.FontBold
'End Property
'
'Public Property Let FontBold(ByVal New_FontBold As Boolean)
'    UserControl.FontBold() = New_FontBold
'    PropertyChanged "FontBold"
'End Property
'
''注意！不要删除或修改下列被注释的行！
''MappingInfo=UserControl,UserControl,-1,FontItalic
'Public Property Get FontItalic() As Boolean
'    FontItalic = UserControl.FontItalic
'End Property
'
'Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
'    UserControl.FontItalic() = New_FontItalic
'    PropertyChanged "FontItalic"
'End Property
'
''注意！不要删除或修改下列被注释的行！
''MappingInfo=UserControl,UserControl,-1,FontName
'Public Property Get FontName() As String
'    FontName = UserControl.FontName
'End Property
'
'Public Property Let FontName(ByVal New_FontName As String)
'    UserControl.FontName() = New_FontName
'    PropertyChanged "FontName"
'End Property
'
''注意！不要删除或修改下列被注释的行！
''MappingInfo=UserControl,UserControl,-1,FontSize
''Public Property Get FontSize() As Single
''    FontSize = UserControl.FontSize
''End Property
''
''Public Property Let FontSize(ByVal New_FontSize As Single)
''    UserControl.FontSize() = New_FontSize
''    PropertyChanged "FontSize"
''End Property
'
''注意！不要删除或修改下列被注释的行！
''MappingInfo=UserControl,UserControl,-1,FontStrikethru
''Public Property Get FontStrikethru() As Boolean
''    FontStrikethru = UserControl.FontStrikethru
''End Property
''
''Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
''    UserControl.FontStrikethru() = New_FontStrikethru
''    PropertyChanged "FontStrikethru"
''End Property
'
''注意！不要删除或修改下列被注释的行！
''MappingInfo=UserControl,UserControl,-1,FontTransparent
'Public Property Get FontTransparent() As Boolean
'    FontTransparent = UserControl.FontTransparent
'End Property
'
'Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
'    UserControl.FontTransparent() = New_FontTransparent
'    PropertyChanged "FontTransparent"
'End Property
'
''注意！不要删除或修改下列被注释的行！
''MappingInfo=UserControl,UserControl,-1,FontUnderline
'Public Property Get FontUnderline() As Boolean
'    FontUnderline = UserControl.FontUnderline
'End Property
'
'Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
'    UserControl.FontUnderline() = New_FontUnderline
'    PropertyChanged "FontUnderline"
'End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "决定是否为该控件分配了唯一的显示上下文。"
    HasDC = UserControl.HasDC
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "返回一个句柄(从 Microsoft Windows)到对象的设备上下文。"
    hDC = UserControl.hDC
End Property

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "返回一个句柄到(from Microsoft Windows)一个对象的窗口。"
    hwnd = UserControl.hwnd
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "返回一个 Microsoft Windows 提供的句柄到一个持久性位图。"
    Set Image = UserControl.Image
End Property


'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "设置一个自定义鼠标图标。"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "返回/设置当鼠标经过对象某一部分时鼠标的指针类型。"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, ByVal Width1 As Variant, ByVal Height1 As Variant, ByVal X2 As Variant, ByVal Y2 As Variant, ByVal Width2 As Variant, ByVal Height2 As Variant, ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "画 Form、PictureBox、或 Printer 对象上的图形文件的内容。"
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "返回/设置控件中显示的图形。"
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

' "Point" 后面的下划线是必须的，
'因为它是 VBA 中的保留字。
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(X As Single, Y As Single) As Long
Attribute Point.VB_Description = "返回一个长整型值，作为 Form 或 PictureBox 对象上指定点的 RGB 颜色值。"
    Point = UserControl.Point(X, Y)
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, ByVal Flags As Variant, ByVal X As Variant, ByVal Y As Variant, ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "显示一个 MDIForm 或 Form 对象上的弹出菜单。"
    UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

' "PSet" 后面的下划线是必须的，
'因为它是 VBA 中的保留字。
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,PSet
Public Sub PSet_(X As Single, Y As Single, Color As Long)
    UserControl.PSet Step(X, Y), Color
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Alpha = m_def_Alpha
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
'    UserControl.FontName = PropBag.ReadProperty("FontName", "")
'    UserControl.FontSize = PropBag.ReadProperty("FontSize", 0)
 '   UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
  '  UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
   ' UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Alpha = PropBag.ReadProperty("Alpha", m_def_Alpha)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, True)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Alpha", m_Alpha, m_def_Alpha)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=1,0,0,100
Public Property Get Alpha() As Byte
Attribute Alpha.VB_Description = "控件透明度（0-100）"
    Alpha = m_Alpha
End Property

Public Property Let Alpha(ByVal New_Alpha As Byte)
    m_Alpha = New_Alpha
    PropertyChanged "Alpha"
    SetTransparentWindow Me.hwnd, New_Alpha
End Property

