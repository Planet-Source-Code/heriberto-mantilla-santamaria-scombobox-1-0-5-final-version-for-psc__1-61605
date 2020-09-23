VERSION 5.00
Begin VB.UserControl CoolList 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   KeyPreview      =   -1  'True
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
   Begin VB.VScrollBar Bar 
      Height          =   915
      Left            =   1200
      Max             =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   1620
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox iScr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   -15
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   2
      Top             =   45
      Width           =   1170
   End
End
Attribute VB_Name = "CoolList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
' CoolList OCX 1.2 (Private UC version)
'
' Carles P.V.
' carles_pv@terra.es
'-------------------------------------------------------------------------------------------
' Last modified: 2003.02.11
'-------------------------------------------------------------------------------------------
Option Explicit

  '* Declares for Unicode support.
 Private Const VER_PLATFORM_WIN32_NT = 2
 
 Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
 End Type
 
 Private mWindowsNT   As Boolean

 Public Enum AlignmentCts
  [AlignLeft] = &H0
  [AlignCenter] = &H2
  [AlignRight] = &H1
 End Enum

 Public Enum AppearanceCts
  [Flat] = &H0
  [3D] = &H1
 End Enum

 Public Enum BorderStyleCts
  [None] = &H0
  [Fixed Single] = &H1
 End Enum

 Public Enum OrderTypeCts
  [Ascendent] = &H0
  [Descendent] = &H1
 End Enum

 Public Enum SelectModeCts
  [Single] = &H0
  [Multiple] = &H1
 End Enum

 Public Enum SelectModeStyleCts
  [Standard] = &H0
  [Dither] = &H1
  [Gradient_V] = &H2
  [Gradient_H] = &H3
  [Box] = &H4
  [Underline] = &H5
  [byPicture] = &H6
 End Enum

 Private Declare Function GetTickCount Lib "kernel32" () As Long
 Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
 Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
 Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
 Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
 Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
 Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
 Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2) As Long
 Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal dx As Long, ByVal dy As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
 Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
 
 Private Type TRIVERTEX
  X     As Long
  Y     As Long
  R     As Integer
  G     As Integer
  B     As Integer
  Alpha As Integer
 End Type

 Private Type RGB
  R As Integer
  G As Integer
  B As Integer
 End Type

 Private Type GRADIENT_RECT
  UpperLeft  As Long
  LowerRight As Long
 End Type

 Private Type RECT2
  x1 As Long
  y1 As Long
  x2 As Long
  y2 As Long
 End Type

 Private Type POINTAPI
  X As Long
  Y As Long
 End Type

 Private Const PS_SOLID As Long = 0

 Private Const GRADIENT_FILL_RECT_H As Long = &H0
 Private Const GRADIENT_FILL_RECT_V As Long = &H1
 
 Private Const DT_LEFT       As Long = &H0
 Private Const DT_VCENTER    As Long = &H4
 Private Const DT_WORDBREAK  As Long = &H10
 Private Const DT_SINGLELINE As Long = &H20
 
 Private Const defSelectBorderColor = &HC56A31
 Private Const defSelectListBorderColor = &H6B2408
 Private Const defShadowColorText = &H80000015

 '-------------------------------------------------------------------------------------------
 ' UserControl constants / types / variables / events
 '-------------------------------------------------------------------------------------------

 Private Type tItem
  Text            As String
  Icon            As Integer
  IconSelected    As Integer
  Color           As Long
  Enabled         As Boolean
  SeparatorLine   As Boolean
  TextShadow      As Boolean
  ToolTipTextItem As String
  MouseIcon       As StdPicture
 End Type

 Private m_List()          As tItem    ' List array of items (Text, icons)
 Private m_Selected()      As Boolean  ' List array of items (Selected/Unselected)
 Private m_nItems          As Integer  ' Number of Items

 Private m_LastBar         As Integer  ' Last scroll bar value
 Private m_LastItem        As Integer  ' Last Selected item
 Private m_LastY           As Single   ' Last Y value [pixels] (prevents item repaint)
 Private m_AnchorItemState As Boolean  ' Anchor item value (multiple selection).
 '  Case extended selection: all selected items
 '  will be set to Anchor selection state.

 Private m_EnsureVisible   As Boolean  ' Ensure visible last m_Selected item (ListIndex)

 Private m_ItemRct()       As RECT2    ' Item rectangle
 Private m_TextRct()       As RECT2    ' Item text rectangle
 Private m_IconPt()        As POINTAPI ' Item icon position
 
 Private m_tmpItemHeight   As Integer  ' Item height [pixels]
 Private m_VisibleRows     As Integer  ' Visible rows in control area
 Private m_Scrolling       As Boolean  ' Scrolling by mouse
 Private m_ScrollingY      As Long     ' Y Scrolling coordinate flag (scroll speed = f(Y))
 Private m_HasFocus        As Boolean  ' Control has focus
 Private m_Resizing        As Boolean  ' Prevent repaints when Resizing

 Private m_pImgList        As Object   ' Will point to ImageList control
 Private m_ILScale         As Integer  ' ImageList parent scale mode

 Private m_Appearance            As AppearanceCts
 Private m_SelectBorderColor     As OLE_COLOR
 Private m_SelectListBorderColor As OLE_COLOR
 Private m_ListGradient          As Boolean
 Private m_ShadowColorText       As OLE_COLOR

 Private m_ColorBack       As Long     ' Back color [Normal]
 Private m_ColorBackSel    As Long     ' Back color [Selected]
 Private m_ColorFont       As Long     ' Font color [Normal]
 Private m_ColorFontSel    As Long     ' Font color [Selected]
 Private m_ColorGradient1  As RGB      ' Gradient color from [Selected]
 Private m_ColorGradient2  As RGB      ' Gradient color to [Selected]
 Private m_ColorBox        As Long     ' Box border color
 
 Private WithEvents m_Font As StdFont  ' Font object
Attribute m_Font.VB_VarHelpID = -1

 Private m_Alignment        As AlignmentCts
 Private m_Apeareance       As AppearanceCts
 Private m_BackNormal       As OLE_COLOR
 Private m_BackSelected     As OLE_COLOR
 Private m_BackSelectedG1   As OLE_COLOR
 Private m_BackSelectedG2   As OLE_COLOR
 Private m_BoxBorder        As OLE_COLOR
 Private m_BoxOffset        As Integer
 Private m_BoxRadius        As Integer
 Private m_Focus            As Boolean
 Private m_FontNormal       As OLE_COLOR
 Private m_FontSelected     As OLE_COLOR
 Private m_HoverSelection   As Boolean
 Private m_ItemHeight       As Integer
 Private m_ItemHeightAuto   As Boolean
 Private m_ItemOffset       As Integer
 Private m_ItemTextLeft     As Integer
 Private m_ListIndex        As Integer
 Private m_OrderType        As OrderTypeCts
 Private m_ScrollBarWidth   As Integer
 Private m_SelectionPicture As Picture
 Private m_SelectMode       As SelectModeCts
 Private m_SelectModeStyle  As SelectModeStyleCts
 Private m_TopIndex         As Integer
 Private m_WordWrap         As Boolean

 Private Const m_def_Appearance = 1
 Private Const m_def_Alignment = DT_LEFT
 Private Const m_def_BackNormal = vbWindowBackground
 Private Const m_def_BackSelected = vbHighlight
 Private Const m_def_BackSelectedG1 = vbHighlight
 Private Const m_def_BackSelectedG2 = vbWindowBackground
 Private Const m_def_BorderStyle = 1
 Private Const m_def_BoxBorder = vbHighlightText
 Private Const m_def_BoxOffset = 1
 Private Const m_def_BoxRadius = 0
 Private Const m_def_Focus = -1
 Private Const m_def_FontNormal = vbWindowText
 Private Const m_def_FontSelected = vbHighlightText
 Private Const m_def_HoverSelection = 0
 Private Const m_def_ItemHeightAuto = -1
 Private Const m_def_ItemOffset = 0
 Private Const m_def_ItemTextLeft = 2
 Private Const m_def_OrderType = 0
 Private Const m_def_ScrollBarWidth = 13
 Private Const m_def_SelectMode = 0
 Private Const m_def_SelectModeStyle = 0
 Private Const m_def_WordWrap = -1

 Public Event Click()
 Public Event DblClick()
 Public Event KeyDown(KeyCode As Integer, Shift As Integer)
 Public Event KeyPress(KeyAscii As Integer)
 Public Event KeyUp(KeyCode As Integer, Shift As Integer)
 Public Event ListIndexChange()
 Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Public Event Scroll()
 Public Event TopIndexChange()
 
'-------------------------------------------------------------------------------------------
' Init/Read/Write properties
'-------------------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
 m_Appearance = m_def_Appearance
 UserControl.BorderStyle = m_def_BorderStyle
 m_ScrollBarWidth = m_def_ScrollBarWidth
 Set iScr.Font = Ambient.Font
 Set m_Font = Ambient.Font
 m_FontNormal = m_def_FontNormal
 m_FontSelected = m_def_FontSelected
 m_BackNormal = m_def_BackNormal
 m_BackSelected = m_def_BackSelected
 m_BackSelectedG1 = m_def_BackSelectedG1
 m_BackSelectedG2 = m_def_BackSelectedG2
 m_BoxBorder = m_def_BoxBorder
 m_BoxOffset = m_def_BoxOffset
 m_BoxRadius = m_def_BoxRadius
 m_Alignment = m_def_Alignment
 m_Focus = m_def_Focus
 m_HoverSelection = m_def_HoverSelection
 m_WordWrap = m_def_WordWrap
 m_ItemHeight = iScr.TextHeight("TextHeight")
 m_ItemHeightAuto = m_def_ItemHeightAuto
 m_ItemOffset = m_def_ItemOffset
 m_ItemTextLeft = m_def_ItemTextLeft
 m_OrderType = m_def_OrderType
 Set m_SelectionPicture = Nothing
 m_SelectMode = m_def_SelectMode
 m_SelectModeStyle = m_def_SelectModeStyle
 m_ListIndex = -1
 m_TopIndex = -1
 m_SelectBorderColor = defSelectBorderColor
 m_SelectListBorderColor = defSelectListBorderColor
 m_ShadowColorText = defShadowColorText
 m_ListGradient = False
 m_VisibleRows = 3
 Call SetColors
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Dim sTmp As String
 
 m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
 UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
 UserControl.Enabled = PropBag.ReadProperty("Enabled", -1)
 m_ScrollBarWidth = PropBag.ReadProperty("ScrollBarWidth", m_def_ScrollBarWidth)
 Bar.Width = PropBag.ReadProperty("ScrollBarWidth", m_def_ScrollBarWidth)
 Set iScr.Font = PropBag.ReadProperty("Font", Ambient.Font)
 Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
 m_FontNormal = PropBag.ReadProperty("FontNormal", m_def_FontNormal)
 m_FontSelected = PropBag.ReadProperty("FontSelected", m_def_FontSelected)
 m_BackNormal = PropBag.ReadProperty("BackNormal", m_def_BackNormal)
 iScr.BackColor = PropBag.ReadProperty("BackNormal", m_def_BackNormal)
 m_BackSelected = PropBag.ReadProperty("BackSelected", m_def_BackSelected)
 m_BackSelectedG1 = PropBag.ReadProperty("BackSelectedG1", m_def_BackSelectedG1)
 m_BackSelectedG2 = PropBag.ReadProperty("BackSelectedG2", m_def_BackSelectedG2)
 m_BoxBorder = PropBag.ReadProperty("BoxBorder", m_def_BoxBorder)
 m_BoxOffset = PropBag.ReadProperty("BoxOffset", m_def_BoxOffset)
 m_BoxRadius = PropBag.ReadProperty("BoxRadius", m_def_BoxRadius)
 m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
 m_Focus = PropBag.ReadProperty("Focus", m_def_Focus)
 m_HoverSelection = PropBag.ReadProperty("HoverSelection", m_def_HoverSelection)
 m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
 m_ItemOffset = PropBag.ReadProperty("ItemOffset", m_def_ItemOffset)
 m_ItemHeightAuto = PropBag.ReadProperty("ItemHeightAuto", m_def_ItemHeightAuto)
 m_ItemTextLeft = PropBag.ReadProperty("ItemTextLeft", m_def_ItemTextLeft)
 m_OrderType = PropBag.ReadProperty("OrderType", m_def_OrderType)
 Set m_SelectionPicture = PropBag.ReadProperty("SelectionPicture", Nothing)
 m_SelectMode = PropBag.ReadProperty("SelectMode", m_def_SelectMode)
 m_SelectModeStyle = PropBag.ReadProperty("SelectModeStyle", m_def_SelectModeStyle)
 iScr.MousePointer = PropBag.ReadProperty("MousePointer", 0)
 Set iScr.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 sTmp = PropBag.ReadProperty("ItemHeight", 0)
 If (sTmp < iScr.TextHeight("")) Then
  m_ItemHeight = iScr.TextHeight("")
 Else
  m_ItemHeight = sTmp
 End If
 m_ListIndex = -1
 m_TopIndex = -1
 Call SetColors
 SelectBorderColor = PropBag.ReadProperty("SelectBorderColor", defSelectBorderColor)
 SelectListBorderColor = PropBag.ReadProperty("SelectListBorderColor", defSelectListBorderColor)
 ShadowColorText = PropBag.ReadProperty("ShadowColorText", defShadowColorText)
 ListGradient = PropBag.ReadProperty("ListGradient", False)
 VisibleRows = PropBag.ReadProperty("VisibleRows", 8)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Appearance", m_Appearance, 1)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
 Call PropBag.WriteProperty("Enabled", UserControl.Enabled, -1)
 Call PropBag.WriteProperty("ScrollBarWidth", m_ScrollBarWidth, m_def_ScrollBarWidth)
 Call PropBag.WriteProperty("Font", iScr.Font, Ambient.Font)
 Call PropBag.WriteProperty("FontNormal", m_FontNormal, m_def_FontNormal)
 Call PropBag.WriteProperty("FontSelected", m_FontSelected, m_def_FontSelected)
 Call PropBag.WriteProperty("BackNormal", m_BackNormal, m_def_BackNormal)
 Call PropBag.WriteProperty("BackSelected", m_BackSelected, m_def_BackSelected)
 Call PropBag.WriteProperty("BackSelectedG1", m_BackSelectedG1, m_def_BackSelectedG1)
 Call PropBag.WriteProperty("BackSelectedG2", m_BackSelectedG2, m_def_BackSelectedG2)
 Call PropBag.WriteProperty("BoxBorder", m_BoxBorder, m_def_BoxBorder)
 Call PropBag.WriteProperty("BoxOffset", m_BoxOffset, m_def_BoxOffset)
 Call PropBag.WriteProperty("BoxRadius", m_BoxRadius, m_def_BoxRadius)
 Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
 Call PropBag.WriteProperty("Focus", m_Focus, m_def_Focus)
 Call PropBag.WriteProperty("HoverSelection", m_HoverSelection, m_def_HoverSelection)
 Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
 Call PropBag.WriteProperty("ItemHeight", m_ItemHeight, 0)
 Call PropBag.WriteProperty("ItemHeightAuto", m_ItemHeightAuto, m_def_ItemHeightAuto)
 Call PropBag.WriteProperty("ItemOffset", m_ItemOffset, m_def_ItemOffset)
 Call PropBag.WriteProperty("ItemTextLeft", m_ItemTextLeft, m_def_ItemTextLeft)
 Call PropBag.WriteProperty("OrderType", m_OrderType, m_def_OrderType)
 Call PropBag.WriteProperty("SelectionPicture", m_SelectionPicture, Nothing)
 Call PropBag.WriteProperty("SelectMode", m_SelectMode, m_def_SelectMode)
 Call PropBag.WriteProperty("SelectModeStyle", m_SelectModeStyle, m_def_SelectModeStyle)
 Call PropBag.WriteProperty("MousePointer", iScr.MousePointer, 0)
 Call PropBag.WriteProperty("MouseIcon", iScr.MouseIcon, Nothing)
 Call PropBag.WriteProperty("SelectBorderColor", m_SelectBorderColor, defSelectBorderColor)
 Call PropBag.WriteProperty("SelectListBorderColor", m_SelectListBorderColor, defSelectListBorderColor)
 Call PropBag.WriteProperty("ShadowColorText", m_ShadowColorText, defShadowColorText)
 Call PropBag.WriteProperty("ListGradient", m_ListGradient, False)
 Call PropBag.WriteProperty("VisibleRows", m_VisibleRows, 8)
End Sub

'-------------------------------------------------------------------------------------------
'-- UserControl initialitation, focus, size, refresh, termination
'-------------------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
 Dim OS As OSVERSIONINFO
 
 '-- Initialize arrays
 ReDim m_List(0)
 ReDim m_Selected(0)
 '-- Initialize position flags
 m_EnsureVisible = -1 ' Ensure visible last selected
 m_LastItem = -1      ' Last selected
 m_LastY = -1         ' Last Y coordinate
 '-- Initialize font object
 Set m_Font = New StdFont
 '* Get the operating system version for text drawing purposes.
 OS.dwOSVersionInfoSize = Len(OS)
 Call GetVersionEx(OS)
 mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
End Sub

Private Sub UserControl_EnterFocus()
 m_HasFocus = -1
 Call DrawFocus(m_ListIndex)
End Sub

Private Sub UserControl_ExitFocus()
 m_HasFocus = 0
 Call DrawItem(m_ListIndex)
End Sub

Private Sub UserControl_Resize()
 '-- Set item height
 If (m_ItemHeightAuto) Then
  m_tmpItemHeight = iScr.TextHeight("")
 Else
  If (m_ItemHeight < iScr.TextHeight("")) Then
   m_tmpItemHeight = iScr.TextHeight("")
  Else
   m_tmpItemHeight = m_ItemHeight
  End If
 End If
 '-- Get visible rows and readjust control height
 'm_VisibleRows = ScaleHeight \ m_tmpItemHeight
 Height = (m_VisibleRows) * m_tmpItemHeight * Screen.TwipsPerPixelX + (Height - ScaleHeight * Screen.TwipsPerPixelY)
 '-- Locate and resize drawing area, calc. rects and readjust scroll bar
 m_Resizing = -1
 Call iScr.Move(0, 0, ScaleWidth - IIf(Bar.Visible, Bar.Width, 0), ScaleHeight)
 With Bar
 On Error Resume Next
  Call .Move(ScaleWidth - .Width, 1, .Width, ScaleHeight - 2)
  .Visible = 0
 End With
 ReDim m_ItemRct(m_VisibleRows - 1)
 ReDim m_TextRct(m_VisibleRows - 1)
 ReDim m_IconPt(m_VisibleRows - 1)
 Call CalculateRects
 Call ReadjustBar
 m_Resizing = 0
End Sub

Private Sub iScr_Paint()
 Dim FocusRect As RECT2
 
 If Not (Ambient.UserMode = True) Then
  Call iScr.Cls
  Select Case m_Alignment
   Case 0: iScr.CurrentX = m_ItemTextLeft + m_ItemOffset
   Case 1: iScr.CurrentX = (ScaleWidth - iScr.TextWidth(Ambient.DisplayName)) - m_ItemOffset
   Case 2: iScr.CurrentX = (ScaleWidth - iScr.TextWidth(Ambient.DisplayName)) * 0.5
  End Select
  iScr.CurrentY = m_ItemOffset
  Call SetTextColor(iScr.hDC, m_ColorFont)
  iScr.Print (Ambient.DisplayName)
  Call SetRect(FocusRect, 0, 0, ScaleWidth, m_tmpItemHeight)
  Call DrawFocusRect(iScr.hDC, FocusRect)
 ElseIf Not (m_Resizing = True) Then
  Call DrawList
 End If
End Sub

Private Sub UserControl_Terminate()
 Erase m_List
 Erase m_Selected
 Set m_pImgList = Nothing
 m_Scrolling = 0
End Sub

'-------------------------------------------------------------------------------------------
'-- ScrollBar
'-------------------------------------------------------------------------------------------
Private Sub Bar_Change()
 If (m_LastBar <> Bar) Then
  m_LastBar = Bar
  m_LastY = -1
  If (txtEdit.Visible = True) Then
   Call txtEdit_LostFocus
  End If
  If (m_ListIndex = m_LastItem) Then
   Call DrawList
  End If
  RaiseEvent Scroll
  RaiseEvent TopIndexChange
 End If
End Sub

Private Sub Bar_Scroll()
 Call Bar_Change
 RaiseEvent Scroll
End Sub

'-------------------------------------------------------------------------------------------
' Scrolling / Events
'-------------------------------------------------------------------------------------------

'-- Click()
Private Sub iScr_Click()
 If (m_ListIndex > -1) Then RaiseEvent Click
End Sub

'-- DblClick()
Private Sub iScr_DblClick()
 If (m_ListIndex > -1) Then RaiseEvent DblClick
End Sub

'-- KeyDown(KeyCode, Shift)
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 If (m_nItems = 0) Or (m_ListIndex = -1) Then
  RaiseEvent KeyDown(KeyCode, Shift)
  Exit Sub
 End If
 Select Case KeyCode
  Case 38 '{Up arrow}
   If (m_ListIndex > 0) Then ListIndex = ListIndex - 1
  Case 40 '{Down arrow}
   If (m_ListIndex < m_nItems - 1) Then ListIndex = ListIndex + 1
  Case 33 '{PageDown}
   If (m_ListIndex > m_VisibleRows) Then
    ListIndex = ListIndex - m_VisibleRows
   Else
    ListIndex = 0
   End If
  Case 34 '{PageUp}
   If (m_ListIndex < m_nItems - m_VisibleRows - 1) Then
    ListIndex = ListIndex + m_VisibleRows
   Else
    ListIndex = m_nItems - 1
   End If
  Case 36 '{Start}
   ListIndex = 0
  Case 35 '{End}
   ListIndex = m_nItems - 1
  Case 32 '{Space} Select/Unselect
   If (m_SelectMode <> 0) And (m_ListIndex > -1) Then
    m_Selected(m_ListIndex) = Not m_Selected(m_ListIndex)
    Call DrawItem(m_ListIndex)
    Call DrawFocus(m_ListIndex)
   End If
  RaiseEvent Click
 End Select
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'-- KeyPress(KeyAscii)
Private Sub UserControl_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
End Sub

'-- KeyPress(KeyCode, Shift)
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'-- MouseDown(Button, Shift, x, y)
Private Sub iScr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim SelectedListIndex As Integer
 
 If (Button = vbRightButton) Then
  RaiseEvent MouseDown(Button, Shift, X, Y)
  Exit Sub
 End If
 SelectedListIndex = Bar + Int(Y / m_tmpItemHeight)
 If (m_List(SelectedListIndex).Enabled = True) And (SelectedListIndex >= 0) And (SelectedListIndex < m_nItems) Then
  Select Case m_SelectMode
   Case 0 ' [Single]
    m_Selected(SelectedListIndex) = -1
   Case 1 ' [Multiple]
    m_Selected(SelectedListIndex) = Not m_Selected(SelectedListIndex)
    m_AnchorItemState = m_Selected(SelectedListIndex)
  End Select
  m_LastY = Y
  ListIndex = SelectedListIndex
 End If
 m_Scrolling = -1
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'-- MouseMove(Button, Shift, x, y)
Private Sub iScr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim SelectedListIndex As Integer

 m_ScrollingY = Y
 If (Y < 0) Then
  Call ScrollUp
  RaiseEvent MouseMove(Button, Shift, X, Y)
  Exit Sub
 End If
 If (Y > ScaleHeight) Then
  Call ScrollDown
  RaiseEvent MouseMove(Button, Shift, X, Y)
  Exit Sub
 End If
 If (m_HoverSelection = True) Or (Button) And (Y \ m_tmpItemHeight <> m_LastY \ m_tmpItemHeight) Then
  SelectedListIndex = Bar + (Y \ m_tmpItemHeight)
  If (SelectedListIndex >= 0) And (SelectedListIndex < m_nItems) Then
   m_Selected(SelectedListIndex) = m_AnchorItemState
   ListIndex = SelectedListIndex
   m_LastY = Y
  End If
 End If
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'-- MouseUp(Button, Shift, x, y)
Private Sub iScr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 m_Scrolling = 0
 m_AnchorItemState = -1
 RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'-------------------------------------------------------------------------------------------
' Methods
'-------------------------------------------------------------------------------------------

'-- SetImageList
Public Sub SetImageList(ImageListControl)
 Set m_pImgList = ImageListControl
On Error Resume Next
 m_ILScale = m_pImgList.Parent.ScaleMode
On Error GoTo 0
 Call iScr_Paint
End Sub

'-- AddItem
'-- 0 , ... , n-1 [n = ListCount]
Public Sub AddItem(ByVal Text As Variant, Optional ByVal Icon As Integer, Optional ByVal IconSelected As Integer, Optional ByVal Color As Long, Optional ByVal Enabled As Boolean = True, Optional ByVal ToolTipTextItem As String = "", Optional ByVal MouseIcon As StdPicture = Nothing, Optional ByVal SeparatorLine As Boolean = False, Optional ByVal TextShadow As Boolean = False)
 m_List(m_nItems).Text = CStr(Text)
 m_List(m_nItems).Icon = Icon
 m_List(m_nItems).IconSelected = IconSelected
 m_List(m_nItems).Color = Color
 m_List(m_nItems).Enabled = Enabled
 m_List(m_nItems).SeparatorLine = SeparatorLine
 Set m_List(m_nItems).MouseIcon = MouseIcon
 m_List(m_nItems).TextShadow = TextShadow
 m_List(m_nItems).ToolTipTextItem = ToolTipTextItem
 m_nItems = m_nItems + 1
 ReDim Preserve m_List(m_nItems)
 ReDim Preserve m_Selected(m_nItems)
 Call ReadjustBar
 If (m_nItems < m_VisibleRows + 1) Then Call DrawItem((m_nItems - 1))
End Sub

'-- InsertItem
Public Sub InsertItem(ByVal Index As Integer, ByVal Text As Variant, Optional ByVal Icon As Integer, Optional ByVal IconSelected As Integer)
 Dim i As Long
 
 If (m_nItems = 0 Or Index > m_nItems) Then Call Err.Raise(381)
 m_nItems = m_nItems + 1
 ReDim Preserve m_List(m_nItems)
 ReDim Preserve m_Selected(m_nItems)
 For i = m_nItems - 1 To Index Step -1
  m_List(i + 1) = m_List(i)
  m_Selected(i + 1) = m_Selected(i)
 Next
 m_List(Index).Text = CStr(Text)
 m_List(Index).Icon = Icon
 m_List(Index).IconSelected = IconSelected
 m_Selected(Index) = 0
 Call ReadjustBar
 m_EnsureVisible = 0
 If (m_ListIndex > -1) And (Index <= m_ListIndex) Then ListIndex = ListIndex + 1
 Call iScr_Paint
End Sub

'-- ModifyItem
Public Sub ModifyItem(ByVal Index As Integer, Optional ByVal Text As Variant = vbEmpty, Optional ByVal Icon As Integer = -1, Optional ByVal IconSelected As Integer = -1)
 If (m_nItems = 0 Or Index > m_nItems) Then Call Err.Raise(381)
 If (Text <> vbEmpty) Then m_List(Index).Text = CStr(Text)
 If (Icon > -1) Then m_List(Index).Icon = Icon
 If (IconSelected > -1) Then m_List(Index).IconSelected = IconSelected
 Call DrawItem(Index)
 Call DrawFocus(m_ListIndex)
End Sub

'-- RemoveItem
Public Sub RemoveItem(ByVal Index As Integer)
 Dim i As Long
 
 If (m_nItems = 0 Or Index > m_nItems - 1) Then Call Err.Raise(381)
 If (Index < m_nItems) Then
  For i = Index To m_nItems - 1
   m_List(i) = m_List(i + 1)
   m_Selected(i) = m_Selected(i + 1)
  Next
 End If
 m_nItems = m_nItems - 1
 ReDim Preserve m_List(m_nItems)
 ReDim Preserve m_Selected(m_nItems)
 Call ReadjustBar
 m_EnsureVisible = 0
 If (Index < m_ListIndex) Then
  If (m_ListIndex > -1) Then ListIndex = ListIndex - 1
 ElseIf (Index = m_ListIndex) Then
  ListIndex = -1
 End If
 If (m_nItems < m_VisibleRows) Then Call iScr.Cls
 Call iScr_Paint
End Sub

'-- FindFirst
Public Function FindFirst(ByVal FindString As String, Optional ByVal StartIndex As Integer = 0, Optional ByVal StartWith As Boolean = False) As Integer
 Dim i As Long
 
 If (m_nItems = 0) Then
  FindFirst = -2
  Exit Function
 End If
 For i = StartIndex To m_nItems
  If (StartWith = True) Then
   If (m_List(i).Text = FindString) Then FindFirst = i: Exit Function
  Else
   If (InStr(1, LCase$(m_List(i).Text), LCase$(FindString)) > 1) Then FindFirst = i: Exit Function
  End If
 Next
 '-- FindString not found
 FindFirst = -1
End Function

'-- Clear
Public Sub Clear()
 '-- Hide scroll bar
 Bar.Visible = 0
 Bar.Max = 0
 '-- Clear and resize drawing area
 Call iScr.Cls
 Call iScr.Move(0, 0, ScaleWidth, ScaleHeight)
 '-- Reset Item arrays
 ReDim m_List(0)
 ReDim m_Selected(0)
 m_nItems = 0
 m_LastItem = -1
 m_ListIndex = -1
 m_TopIndex = -1
End Sub

'-- Order
Public Sub Order()
 Dim i0     As Long, i1     As Long
 Dim i2     As Long, d      As Long
 Dim xItem  As tItem, bDesc As Boolean

 If (m_nItems > 1) Then
  i0 = 0
  bDesc = (m_OrderType = [Descendent])
  If (m_SelectMode = [Single]) Then
   If (m_ListIndex > -1) Then m_Selected(m_ListIndex) = 0
  End If
  Do
   d = d * 3 + 1
  Loop Until (d > m_nItems)
  Do
   d = d \ 3
   For i1 = d + i0 To m_nItems + i0 - 1
    xItem = m_List(i1)
    i2 = i1
    Do While ((m_List(i2 - d).Text > xItem.Text) Xor bDesc)
     m_List(i2) = m_List(i2 - d)
     i2 = i2 - d
     If (i2 - d < i0) Then Exit Do
    Loop
    m_List(i2) = xItem
   Next
  Loop Until (d = 1)
  ListIndex = -1
  Bar = 0
  '-- Unselect all and refresh
  ReDim m_Selected(0 To m_nItems)
  Call iScr_Paint
 End If
End Sub

'-------------------------------------------------------------------------------------------
'-- Draw List / Item / Focus
'-------------------------------------------------------------------------------------------

'-- DrawList
Private Sub DrawList()
 Dim i As Long

 iScr.Cls
 If (UBound(m_List) > 0) Then
  '-- Draw visible rows
  Call DrawGrad
  For i = Bar To Bar + m_VisibleRows - 1
   Call DrawItem(i)
  Next
  '-- Draw focus
  Call DrawFocus(m_ListIndex)
 End If
 Call APIRectangle(iScr.hDC, 0, 0, iScr.ScaleWidth - 1, iScr.ScaleHeight - 1, m_SelectBorderColor)
End Sub

'-- DrawItem
Private Sub DrawItem(ByVal Index As Integer)
 Dim nRctIndex As Integer, FontC As Long

 '-- Item out of area?
 If (Index < Bar) Or (Index > Bar + m_VisibleRows - 1) Then Exit Sub
 If (Index > UBound(m_List) - 1) Then Exit Sub
 iScr.FontUnderline = 0
 nRctIndex = Index - Bar
On Error Resume Next
 If Not (m_Selected(Index - 1) = True) And (m_Selected(Index) = False) And (m_List(Index).SeparatorLine = True) Then
  Call APIRectangle(iScr.hDC, 5, m_ItemRct(nRctIndex).y1, iScr.ScaleWidth, 0, m_ShadowColorText)
 ElseIf (m_List(Index).SeparatorLine = True) And (m_List(Index - 1).Enabled = False) Then
  Call APIRectangle(iScr.hDC, 5, m_ItemRct(nRctIndex).y1, iScr.ScaleWidth, 0, m_ShadowColorText)
 End If
 '-- Draw m_Selected Item
 If (m_Selected(Index) = True) And (m_List(Index).Enabled = True) Then
  '-- Draw back area
  Select Case m_SelectModeStyle
   Case 0 '[Standard]
    Call DrawBack(iScr.hDC, m_ItemRct(nRctIndex), m_ColorBackSel, True)
    Call SetTextColor(iScr.hDC, m_ColorFontSel)
   Case 1 '[Dither] *(Effect will be applied after drawing icon)
    Call DrawBack(iScr.hDC, m_ItemRct(nRctIndex), m_ColorBack, True)
    Call SetTextColor(iScr.hDC, m_ColorFontSel)
   Case 2 '[Gradient_V]
    Call DrawBackGrad(iScr.hDC, m_ItemRct(nRctIndex), m_ColorGradient1, m_ColorGradient2, GRADIENT_FILL_RECT_V)
    Call SetTextColor(iScr.hDC, m_ColorFontSel)
   Case 3 '[Gradient_H]
    Call DrawBackGrad(iScr.hDC, m_ItemRct(nRctIndex), m_ColorGradient1, m_ColorGradient2, GRADIENT_FILL_RECT_H)
    Call SetTextColor(iScr.hDC, m_ColorFontSel)
   Case 4 '[Box]
    Call DrawBack(iScr.hDC, m_ItemRct(nRctIndex), m_ColorBack, True)
    Call DrawBox(iScr.hDC, m_ItemRct(nRctIndex), m_BoxOffset, m_BoxRadius, m_ColorBackSel, m_ColorBox)
    Call SetTextColor(iScr.hDC, m_ColorFontSel)
   Case 5 '[Underline]
    Call DrawBack(iScr.hDC, m_ItemRct(nRctIndex), m_ColorBack)
    Call SetTextColor(iScr.hDC, m_ColorFontSel)
    iScr.FontUnderline = -1
   Case 6 '[byPicture]
    If Not (SelectionPicture Is Nothing) Then
     Call iScr.PaintPicture(SelectionPicture, 0, m_ItemRct(nRctIndex).y1, m_ItemRct(nRctIndex).x2, m_tmpItemHeight)
    Else
     Call DrawBack(iScr.hDC, m_ItemRct(nRctIndex), m_ColorBackSel, True)
    End If
    Call SetTextColor(iScr.hDC, m_ColorFontSel)
  End Select
  '-- Draw icon
  If (Not m_pImgList Is Nothing) Then
  On Error Resume Next
   If (m_WordWrap = True) Then
    Call m_pImgList.ListImages(m_List(Index).IconSelected).Draw(iScr.hDC, ScaleX(m_ItemOffset + 1, vbPixels, m_ILScale), ScaleY(m_ItemRct(nRctIndex).y1 + m_ItemOffset, vbPixels, m_ILScale), 1)
   Else
    Call m_pImgList.ListImages(m_List(Index).IconSelected).Draw(iScr.hDC, ScaleX(m_ItemOffset + 1, vbPixels, m_ILScale), ScaleY(m_ItemRct(nRctIndex).y1 + (m_tmpItemHeight - m_pImgList.ImageHeight) * 0.5, vbPixels, m_ILScale), 1)
   End If
  On Error GoTo 0
  End If
  '-- Apply dither effect (*)
  If (m_SelectModeStyle = 1) Then Call DrawDither(iScr.hDC, m_ItemRct(nRctIndex), m_ColorBackSel)
 Else
  '-- Draw back area
  Call SetTextColor(iScr.hDC, m_List(Index).Color)
  '-- Draw icon
  If (Not m_pImgList Is Nothing) Then
  On Error Resume Next
   If (m_List(Index).Enabled = True) Then
    If (m_WordWrap = True) Then
     Call m_pImgList.ListImages(m_List(Index).IconSelected).Draw(iScr.hDC, ScaleX(m_ItemOffset + 1, vbPixels, m_ILScale), ScaleY(m_ItemRct(nRctIndex).y1 + m_ItemOffset, vbPixels, m_ILScale), 1)
    Else
     Call m_pImgList.ListImages(m_List(Index).IconSelected).Draw(iScr.hDC, ScaleX(m_ItemOffset + 1, vbPixels, m_ILScale), ScaleY(m_ItemRct(nRctIndex).y1 + (m_tmpItemHeight - m_pImgList.ImageHeight) * 0.5, vbPixels, m_ILScale), 1)
    End If
   Else
    Call RenderIconGrayscale(iScr.hDC, m_pImgList.ListImages(Index + 1).ExtractIcon.Handle, m_ItemOffset + 1, m_ItemRct(nRctIndex).y1 + m_ItemOffset)
   End If
  On Error GoTo 0
  End If
 End If
 If (m_Selected(Index) = True) And (m_List(Index).Enabled = True) Then
  FontC = m_ColorFontSel
 Else
  FontC = m_List(Index).Color
 End If
 '-- Draw text...
 If (m_WordWrap = True) Then
  If (m_List(Index).TextShadow = True) Then
   Call SetTextColor(iScr.hDC, m_ShadowColorText)
   m_TextRct(nRctIndex).x1 = m_TextRct(nRctIndex).x1 + 2: m_TextRct(nRctIndex).x2 = m_TextRct(nRctIndex).x2 + 1
   If (mWindowsNT = True) Then
    Call DrawTextW(iScr.hDC, StrPtr(m_List(Index).Text), Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
   Else
    Call DrawTextA(iScr.hDC, m_List(Index).Text, Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
   End If
   m_TextRct(nRctIndex).x1 = m_TextRct(nRctIndex).x1 - 2: m_TextRct(nRctIndex).x2 = m_TextRct(nRctIndex).x2 - 1
  End If
  Call SetTextColor(iScr.hDC, FontC)
  If (mWindowsNT = True) Then
   Call DrawTextW(iScr.hDC, StrPtr(m_List(Index).Text), Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
  Else
   Call DrawTextA(iScr.hDC, m_List(Index).Text, Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
  End If
 Else
  If (m_List(Index).TextShadow = True) Then
   Call SetTextColor(iScr.hDC, m_ShadowColorText)
   m_TextRct(nRctIndex).x1 = m_TextRct(nRctIndex).x1 + 1: m_TextRct(nRctIndex).x2 = m_TextRct(nRctIndex).x2 + 1
   If (mWindowsNT = True) Then
    Call DrawTextW(iScr.hDC, StrPtr(m_List(Index).Text), Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
   Else
    Call DrawTextA(iScr.hDC, m_List(Index).Text, Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
   End If
   m_TextRct(nRctIndex).x1 = m_TextRct(nRctIndex).x1 - 1: m_TextRct(nRctIndex).x2 = m_TextRct(nRctIndex).x2 - 1
  End If
  Call SetTextColor(iScr.hDC, FontC)
  If (mWindowsNT = True) Then
   Call DrawTextW(iScr.hDC, StrPtr(m_List(Index).Text), Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
  Else
   Call DrawTextA(iScr.hDC, m_List(Index).Text, Len(m_List(Index).Text), m_TextRct(nRctIndex), m_Alignment Or DT_WORDBREAK)
  End If
 End If
End Sub

'-- DrawFocus
Private Sub DrawFocus(ByVal Index As Integer)
 If Not (m_Focus = True) Or Not (m_HasFocus = True) Then Exit Sub
 '-- Item out of area ?
 If (Index < Bar) Or (Index > Bar + m_VisibleRows - 1) Then Exit Sub
 '-- Draw it
 Call SetTextColor(iScr.hDC, m_ColorFont)
 Call DrawFocusRect(iScr.hDC, m_ItemRct(Index - Bar))
End Sub

Private Sub DrawBack(ByVal hDC As Long, pRect As RECT2, ByVal Color As Long, Optional ByVal Selected As Boolean = False)
 Dim hBrush As Long

On Error Resume Next
 hBrush = CreateSolidBrush(Color)
 Call FillRect(hDC, pRect, hBrush)
 Call DeleteObject(hBrush)
 If (Selected = True) Then
  hBrush = CreateSolidBrush(m_SelectListBorderColor)
  Call FrameRect(hDC, pRect, hBrush)
  Call DeleteObject(hBrush)
 End If
End Sub

Private Sub DrawDither(ByVal hDC As Long, pRect As RECT2, ByVal Color As Long)
 Dim hBrush As Long

 hBrush = SelectObject(hDC, CreateSolidBrush(Color))
 Call PatBlt(hDC, pRect.x1, pRect.y1, pRect.x2 - pRect.x1, pRect.y2 - pRect.y1, &HA000C9)
 Call DeleteObject(SelectObject(hDC, hBrush))
End Sub

Private Sub DrawBackGrad(ByVal hDC As Long, pRect As RECT2, Color1 As RGB, Color2 As RGB, ByVal Direction As Long)
 Dim v(1) As TRIVERTEX
 Dim GRct As GRADIENT_RECT

 '-- from
 With v(0)
  .X = pRect.x1
  .Y = pRect.y1
  .R = Color1.R
  .G = Color1.G
  .B = Color1.B
  .Alpha = 0
 End With
 '-- to
 With v(1)
  .X = pRect.x2
  .Y = pRect.y2
  .R = Color2.R
  .G = Color2.G
  .B = Color2.B
  .Alpha = 0
 End With
 GRct.UpperLeft = 0
 GRct.LowerRight = 1
 Call GradientFillRect(hDC, v(0), 2, GRct, 1, Direction)
End Sub

Private Sub DrawBox(ByVal hDC As Long, pRect As RECT2, ByVal Offset As Long, ByVal Radius As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long)
 Dim hPen As Long, hBrush As Long

 hPen = SelectObject(hDC, CreatePen(PS_SOLID, 1, ColorBorder))
 hBrush = SelectObject(hDC, CreateSolidBrush(ColorFill))
 Call InflateRect(pRect, -Offset, -Offset)
 Call RoundRect(hDC, pRect.x1, pRect.y1, pRect.x2, pRect.y2, Radius, Radius)
 Call InflateRect(pRect, Offset, Offset)
 Call DeleteObject(SelectObject(hDC, hPen))
 Call DeleteObject(SelectObject(hDC, hBrush))
End Sub

Private Sub ReadjustBar()
 If (m_nItems > m_VisibleRows) Then
  If Not (Bar.Visible = True) Then
   '-- Show scroll bar
   Bar.Visible = -1
   Call Bar.Refresh
   Bar.LargeChange = IIf(m_VisibleRows = 0, 1, m_VisibleRows)
   '-- Update item rects. right margin
   Call RigthOffsetRects(Bar.Width)
   '-- Repaint control area
   Call iScr_Paint
  End If
 Else
  '-- Hide scroll bar
  Bar.Visible = 0
  '-- Update item rects. right margin
  Call RigthOffsetRects(0)
 End If
 '-- Update Bar max value
 Bar.Max = m_nItems - m_VisibleRows
End Sub

Private Sub CalculateRects()
 Dim i As Long

 For i = 0 To m_VisibleRows - 1
  Call SetRect(m_ItemRct(i), 0, i * m_tmpItemHeight, ScaleWidth, i * m_tmpItemHeight + m_tmpItemHeight)
  Call SetRect(m_TextRct(i), m_ItemOffset + m_ItemTextLeft, i * m_tmpItemHeight + m_ItemOffset, ScaleWidth - m_ItemOffset, i * m_tmpItemHeight + m_tmpItemHeight - m_ItemOffset)
  m_IconPt(i).X = m_ItemOffset
  m_IconPt(i).Y = m_ItemOffset
 Next
End Sub

Private Sub RigthOffsetRects(ByVal Offset As Long)
 Dim i As Long

 For i = 0 To m_VisibleRows - 1
  m_ItemRct(i).x2 = ScaleWidth - Offset
  m_TextRct(i).x2 = ScaleWidth - m_ItemOffset - Offset
 Next
End Sub

'-------------------------------------------------------------------------------------------
' Scroll Up/Down by mouse / multiple select
'-------------------------------------------------------------------------------------------

'-- ScrollUp
Private Sub ScrollUp()
 Dim t As Long ' Timer counter
 Dim d As Long ' Scrolling delay

 d = 500 + 20 * m_ScrollingY
 If (d < 40) Then d = 40
 '-- Scroll while MouseDown and mouse pos. < "Control top"
 Do While (m_Scrolling = True) And (m_ScrollingY < 0)
  If (GetTickCount - t > d) Then
   t = GetTickCount
   If (m_ListIndex > 0) Then
    If (m_SelectMode = [Multiple]) Then
     m_Selected(m_ListIndex - 1) = m_AnchorItemState
    End If
    ListIndex = ListIndex - 1
   End If
  End If
  DoEvents
 Loop
End Sub

'-- ScrollDown
Private Sub ScrollDown()
 Dim t As Long ' Timer counter
 Dim d As Long ' Scrolling delay

 d = 500 - 20 * (m_ScrollingY - ScaleHeight - 1)
 If (d < 40) Then d = 40
 '-- Scroll while MouseDown and mouse pos. > "Control bottom"
 Do While (m_Scrolling = True) And (m_ScrollingY > ScaleHeight - 1)
  If (GetTickCount - t > d) Then
   t = GetTickCount
   If (m_ListIndex < m_nItems - 1) Then
    If (m_SelectMode = [Multiple]) Then
     m_Selected(m_ListIndex + 1) = m_AnchorItemState
    End If
    ListIndex = ListIndex + 1
   End If
  End If
  DoEvents
 Loop
End Sub

'-------------------------------------------------------------------------------------------
' Colors
'-------------------------------------------------------------------------------------------

'-- SetColors
Private Sub SetColors()
 m_ColorBack = GetLngColor(m_BackNormal)
 m_ColorBackSel = GetLngColor(m_BackSelected)
 m_ColorGradient1 = GetRGBColors(GetLngColor(m_BackSelectedG1))
 m_ColorGradient2 = GetRGBColors(GetLngColor(m_BackSelectedG2))
 m_ColorBox = GetLngColor(m_BoxBorder)
 m_ColorFont = GetLngColor(m_FontNormal)
 m_ColorFontSel = GetLngColor(m_FontSelected)
End Sub

Private Function GetLngColor(ByVal Color As Long) As Long
 If (Color And &H80000000) Then
  GetLngColor = GetSysColor(Color And &H7FFFFFFF)
 Else
  GetLngColor = Color
 End If
End Function

Private Function GetRGBColors(ByVal Color As Long) As RGB
 Dim HexColor As String

 HexColor = String$(6 - Len(Hex(Color)), "0") & Hex$(Color)
 GetRGBColors.R = "&H" & Mid$(HexColor, 5, 2) & "00"
 GetRGBColors.G = "&H" & Mid$(HexColor, 3, 2) & "00"
 GetRGBColors.B = "&H" & Mid$(HexColor, 1, 2) & "00"
End Function

'-------------------------------------------------------------------------------------------
' Properties
'-------------------------------------------------------------------------------------------

'-- Alignment
Public Property Get Alignment() As AlignmentCts
 Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentCts)
 m_Alignment = New_Alignment
 Call iScr_Paint
End Property

'-- Appearance
Public Property Get Appearance() As AppearanceCts
 Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceCts)
 m_Appearance = New_Appearance
End Property

'-- BackNormal
Public Property Get BackNormal() As OLE_COLOR
 BackNormal = m_BackNormal
End Property

Public Property Let BackNormal(ByVal New_BackNormal As OLE_COLOR)
 m_BackNormal = New_BackNormal
 m_ColorBack = GetLngColor(m_BackNormal)
 iScr.BackColor = m_ColorBack
 Call iScr_Paint
End Property

'-- BackSelected
Public Property Get BackSelected() As OLE_COLOR
 BackSelected = m_BackSelected
End Property

Public Property Let BackSelected(ByVal New_BackSelected As OLE_COLOR)
 m_BackSelected = New_BackSelected
 m_ColorBackSel = GetLngColor(m_BackSelected)
 Call iScr_Paint
End Property

'-- BackSelectedG1
Public Property Get BackSelectedG1() As OLE_COLOR
 BackSelectedG1 = m_BackSelectedG1
End Property

Public Property Let BackSelectedG1(ByVal New_BackSelectedG1 As OLE_COLOR)
 m_BackSelectedG1 = New_BackSelectedG1
 m_ColorGradient1 = GetRGBColors(GetLngColor(m_BackSelectedG1))
 Call iScr_Paint
End Property

'-- BackSelectedG2
Public Property Get BackSelectedG2() As OLE_COLOR
 BackSelectedG2 = m_BackSelectedG2
End Property

Public Property Let BackSelectedG2(ByVal New_BackSelectedG2 As OLE_COLOR)
 m_BackSelectedG2 = New_BackSelectedG2
 m_ColorGradient2 = GetRGBColors(GetLngColor(m_BackSelectedG2))
 Call iScr_Paint
End Property

'-- BorderStyle
Public Property Get BorderStyle() As BorderStyleCts
 BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleCts)
 UserControl.BorderStyle() = New_BorderStyle
End Property

'-- BoxBorder
Public Property Get BoxBorder() As OLE_COLOR
 BoxBorder = m_BoxBorder
End Property

Public Property Let BoxBorder(ByVal New_BoxBorder As OLE_COLOR)
 m_BoxBorder = New_BoxBorder
 m_ColorBox = GetLngColor(m_BoxBorder)
 Call iScr_Paint
End Property

'-- BoxOffset
Public Property Get BoxOffset() As Integer
 BoxOffset = m_BoxOffset
End Property

Public Property Let BoxOffset(ByVal New_BoxOffset As Integer)
 If (New_BoxOffset <= m_tmpItemHeight * 0.5) Then
  m_BoxOffset = New_BoxOffset
 End If
 Call iScr_Paint
End Property

'-- BoxRadius
Public Property Get BoxRadius() As Integer
 BoxRadius = m_BoxRadius
End Property

Public Property Let BoxRadius(ByVal New_BoxRadius As Integer)
 m_BoxRadius = New_BoxRadius
 Call iScr_Paint
End Property

'-- Enabled
Public Property Get Enabled() As Boolean
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 UserControl.Enabled() = New_Enabled
 Bar.Enabled = New_Enabled
 iScr.Enabled = New_Enabled
End Property

'-- Focus
Public Property Get Focus() As Boolean
 Focus = m_Focus
End Property

Public Property Let Focus(ByVal New_Focus As Boolean)
 m_Focus = New_Focus
 If (New_Focus) Then
  Call DrawFocus(m_ListIndex)
 Else
  Call DrawItem(m_ListIndex)
 End If
End Property

'-- Font
Public Property Get Font() As Font
 Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
 With m_Font
  .Name = New_Font.Name
  .Size = New_Font.Size
  .Bold = New_Font.Bold
  .Italic = New_Font.Italic
  .Underline = New_Font.Underline
  .Strikethrough = New_Font.Strikethrough
 End With
 Call iScr_Paint
End Property

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
 Set iScr.Font = m_Font
 Call UserControl_Resize
End Sub

'-- FontNormal
Public Property Get FontNormal() As OLE_COLOR
 FontNormal = m_FontNormal
End Property

Public Property Let FontNormal(ByVal New_FontNormal As OLE_COLOR)
 m_FontNormal = New_FontNormal
 m_ColorFont = GetLngColor(m_FontNormal)
 Call SetTextColor(iScr.hDC, m_ColorFont)
 Call iScr_Paint
End Property

'-- FontSelected
Public Property Get FontSelected() As OLE_COLOR
 FontSelected = m_FontSelected
End Property

Public Property Let FontSelected(ByVal New_FontSelected As OLE_COLOR)
 m_FontSelected = New_FontSelected
 m_ColorFontSel = GetLngColor(m_FontSelected)
 Call iScr_Paint
End Property

'-- HoverSelection
Public Property Get HoverSelection() As Boolean
 HoverSelection = m_HoverSelection
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
 m_HoverSelection = New_HoverSelection
 Call DrawItem(m_ListIndex)
 Call DrawFocus(m_ListIndex)
End Property

'-- hWnd
Public Property Get hWnd() As Long
 hWnd = UserControl.hWnd
End Property

'-- ItemHeight
Public Property Get ItemHeight() As Integer
 ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Integer)
 m_ItemHeight = New_ItemHeight
 Call UserControl_Resize
 Call iScr_Paint
End Property

'-- ItemHeightAuto
Public Property Get ItemHeightAuto() As Boolean
 ItemHeightAuto = m_ItemHeightAuto
End Property

Public Property Let ItemHeightAuto(ByVal New_ItemHeightAuto As Boolean)
 m_ItemHeightAuto = New_ItemHeightAuto
 Call UserControl_Resize
 Call iScr_Paint
End Property

'-- ItemOffset
Public Property Get ItemOffset() As Integer
 ItemOffset = m_ItemOffset
End Property

Public Property Let ItemOffset(ByVal New_ItemOffset As Integer)
 If (New_ItemOffset <= m_tmpItemHeight) Then
  m_ItemOffset = New_ItemOffset
 End If
 Call CalculateRects
 If (Bar.Visible = True) Then Call RigthOffsetRects(Bar.Width)
 Call iScr_Paint
End Property

'-- ItemTextLeft
Public Property Get ItemTextLeft() As Integer
 ItemTextLeft = m_ItemTextLeft
End Property

Public Property Let ItemTextLeft(ByVal New_ItemTextLeft As Integer)
 m_ItemTextLeft = New_ItemTextLeft
 Call CalculateRects
 If (Bar.Visible = True) Then Call RigthOffsetRects(Bar.Width)
 Call iScr_Paint
End Property

'-- <ListCount>
Public Property Get ListCount() As Integer
 ListCount = m_nItems
End Property

'-- ListIndex
Public Property Get ListIndex() As Integer
 ListIndex = m_ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
 If (New_ListIndex < -1) Or (New_ListIndex > m_nItems - 1) Then Call Err.Raise(380)
 If (txtEdit.Visible = True) Then Call txtEdit_LostFocus
 If (New_ListIndex < 0 Or m_nItems = 0) Then
  m_ListIndex = -1
  m_LastY = -1
 Else
  m_ListIndex = New_ListIndex
 End If
 '-- Unselect last / Select actual [Single selection mode]
 If (m_SelectMode = [Single]) Then
  If (m_LastItem > -1) Then m_Selected(m_LastItem) = 0
  If (m_ListIndex > -1) Then m_Selected(m_ListIndex) = -1
 End If
 '-- Draw last (delete Focus) ...
 Call Refresh
 Call DrawItem(m_LastItem)
 m_LastItem = m_ListIndex
 '-- ... and draw actual (draw Focus)
 Call DrawItem(m_ListIndex)
 Call DrawFocus(m_ListIndex)
 '-- Ensure visible actual Selected item
 If (m_EnsureVisible = True) Then
  If (m_ListIndex < Bar) And (m_ListIndex > -1) Then
   Bar = m_ListIndex
  ElseIf (m_ListIndex > Bar + m_VisibleRows - 1) Then
   Bar = m_ListIndex - m_VisibleRows + 1
  End If
 Else
  m_EnsureVisible = -1
 End If
 RaiseEvent ListIndexChange
End Property

'-- MouseIcon
Public Property Get MouseIcon() As Picture
 Set MouseIcon = iScr.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
 Set iScr.MouseIcon = New_MouseIcon
End Property

'-- MousePointer
Public Property Get MousePointer() As MousePointerConstants
 MousePointer = iScr.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
 iScr.MousePointer() = New_MousePointer
End Property

'-- OrderType
Public Property Get OrderType() As OrderTypeCts
 OrderType = m_OrderType
End Property

Public Property Let OrderType(ByVal New_OrderType As OrderTypeCts)
 m_OrderType = New_OrderType
End Property

'-- ScrollBarWidth
Public Property Get ScrollBarWidth() As Integer
 ScrollBarWidth = m_ScrollBarWidth
End Property

Public Property Let ScrollBarWidth(ByVal New_ScrollBarWidth As Integer)
 '-- Check Min value width...
 If (New_ScrollBarWidth < 9) Then
  m_ScrollBarWidth = 9
  Bar.Width = 9
  '-- Check Max value width...
 ElseIf (New_ScrollBarWidth > ScaleWidth * 0.5) Then
  m_ScrollBarWidth = ScaleWidth * 0.5
  Bar.Width = ScaleWidth * 0.5
  '-- Set new value...
 Else
  m_ScrollBarWidth = New_ScrollBarWidth
  Bar.Width = New_ScrollBarWidth
 End If
 Bar.Visible = 0
 Call ReadjustBar
 Call UserControl_Resize
End Property

'-- <SelectedCount>
Public Property Get SelectedCount() As Integer
 Dim i As Long

 SelectedCount = 0
 For i = 0 To m_nItems
  If (m_Selected(i)) Then SelectedCount = SelectedCount + 1
 Next
End Property

'-- SelectionPicture
Public Property Get SelectionPicture() As Picture
 Set SelectionPicture = m_SelectionPicture
End Property

Public Property Set SelectionPicture(ByVal New_SelectionPicture As Picture)
 Set m_SelectionPicture = New_SelectionPicture
 Call iScr_Paint
End Property

'-- SelectMode
Public Property Get SelectMode() As SelectModeCts
 SelectMode = m_SelectMode
End Property

Public Property Let SelectMode(ByVal New_SelectMode As SelectModeCts)
 Dim i As Long
 
 m_SelectMode = New_SelectMode
 If (Ambient.UserMode = True) Then
  If (New_SelectMode = [Single]) Then
   '-- Unselect all and select actual
   If (m_ListIndex > -1) Then
    For i = LBound(m_List) To m_nItems
     If (i <> m_ListIndex) Then m_Selected(i) = 0
    Next
    m_Selected(m_ListIndex) = -1
    Call DrawItem(m_ListIndex)
    Call DrawFocus(m_ListIndex)
   End If
  End If
 End If
 Call ReadjustBar
 Call iScr_Paint
End Property

'-- SelectModeStyle
Public Property Get SelectModeStyle() As SelectModeStyleCts
 SelectModeStyle = m_SelectModeStyle
End Property

Public Property Let SelectModeStyle(ByVal New_SelectModeStyle As SelectModeStyleCts)
 m_SelectModeStyle = New_SelectModeStyle
 Call iScr_Paint
End Property

'-- TopIndex
Public Property Get TopIndex() As Integer
 TopIndex = Bar
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)
 If (New_TopIndex < 0) Or (New_TopIndex > m_nItems - m_VisibleRows) Then
  Exit Property
  'Call Err.Raise(380)
 End If
 m_TopIndex = New_TopIndex
 Bar = New_TopIndex
 RaiseEvent TopIndexChange
End Property

'-- WordWrap
Public Property Get WordWrap() As Boolean
 WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
 m_WordWrap = New_WordWrap
 Call iScr_Paint
End Property

'Last revised: 02/07/02
'-------------------------------------------------------------------------------------------
' Some methods passed to R/W properties:
'
' GetItem i    GetIcon i    GetIconSelected i    IsSelected i
' to           to           to                   to
' ItemText(i)  ItemIcon(i)  ItemIconSelected(i)  ItemSelected(i)
'
' Or use ModifyItem to change all item parameters at time


'-- ItemText
Public Property Get ItemText(ByVal Index As Integer) As String
 If (m_nItems = 0 Or Index > m_nItems) Then Call Err.Raise(381)
 ItemText = m_List(Index).Text
End Property

Public Property Let ItemText(ByVal Index As Integer, ByVal Data As String)
 If (m_nItems = 0 Or Index > m_nItems) Then Call Err.Raise(381)
 m_List(Index).Text = CStr(Data)
 Call DrawItem(Index)
 Call DrawFocus(m_ListIndex)
End Property

'-- ItemIcon
Public Property Get ItemIcon(ByVal Index As Integer) As Integer
 If (m_nItems = 0) Or (Index > m_nItems) Then Call Err.Raise(381)
 ItemIcon = m_List(Index).Icon
End Property

Public Property Let ItemIcon(ByVal Index As Integer, ByVal Data As Integer)
 If (m_nItems = 0) Or (Index > m_nItems) Then Call Err.Raise(381)
 m_List(Index).Icon = Data
 Call DrawItem(Index)
 Call DrawFocus(m_ListIndex)
End Property

Public Property Get ItemPicture(ByVal Index As Integer)
 Set ItemPicture = m_pImgList.ListImages(Index).ExtractIcon
End Property

'-- ItemIconSelected
Public Property Get ItemIconSelected(ByVal Index As Integer) As Integer
 If (m_nItems = 0 Or Index > m_nItems) Then Call Err.Raise(381)
 ItemIconSelected = m_List(Index).IconSelected
End Property

Public Property Let ItemIconSelected(ByVal Index As Integer, ByVal Data As Integer)
 If (m_nItems = 0 Or Index > m_nItems) Then Call Err.Raise(381)
 m_List(Index).IconSelected = Data
 Call DrawItem(Index)
 Call DrawFocus(m_ListIndex)
End Property

'-- ItemSelected
Public Property Get ItemSelected(ByVal Index As Integer) As Boolean
 If (m_nItems = 0 Or Index > m_nItems) Then Call Err.Raise(381)
 ItemSelected = m_Selected(Index)
End Property

Public Property Let ItemSelected(ByVal Index As Integer, ByVal Data As Boolean)
 If (m_nItems = 0) Or (Index > m_nItems) Then Call Err.Raise(381)
 Select Case Data
  Case -1
   If (m_SelectMode = [Single]) Then
    ListIndex = Index
   Else
    m_Selected(Index) = -1
    Call DrawItem(Index)
    If (Index = m_ListIndex) Then Call DrawFocus(Index)
   End If
  Case 0
   If Not (m_SelectMode = [Single]) Then
    m_Selected(Index) = 0
    Call DrawItem(Index)
    If (Index = m_ListIndex) Then Call DrawFocus(Index)
   End If
 End Select
End Property

'Editing item...
'-------------------------------------------------------------------------------------------
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
 ' WordWrap mode enabled:
 ' [Control]+[Return] = new line
 ' [Return]           = update text
 ' WordWrap mode disabled:
 ' [Return]           = update text

 '-- Enabled new line in WordWrap mode
 If (m_WordWrap = True) Then
  If (KeyAscii = 13) Then
   m_List(m_ListIndex).Text = txtEdit
   Call txtEdit_LostFocus
  End If
  '-- Don't allow new line in disabled WordWrap mode
 Else
  If (KeyAscii = 13) Or (KeyAscii = 10) Then
   m_List(m_ListIndex).Text = txtEdit
   Call txtEdit_LostFocus
  End If
 End If
 '-- Cancel edition
 If (KeyAscii = 27) Then Call txtEdit_LostFocus
End Sub

Private Sub txtEdit_LostFocus()
 '-- Hide edit TextBox and let ListBox keyboard control
 txtEdit.Visible = 0
 KeyPreview = -1
End Sub

Public Sub StartEdit()
 '-- Item is selected...
 If (m_ListIndex > -1) Then
  '-- Let TextBox keyboard control
  KeyPreview = 0
  With txtEdit
   '-- Get TextBox item font properties
   Set .Font = m_Font
   If (m_Selected(m_ListIndex)) And (m_SelectModeStyle <> [Underline]) Then
    .BackColor = m_ColorBackSel
    .ForeColor = m_ColorFontSel
   Else
    .BackColor = m_ColorBack
    .ForeColor = m_ColorFont
   End If
   '-- Set alignment. Locate and resize TextBox
   If (m_WordWrap = True) Then
    .Alignment = Choose(m_Alignment + 1, 0, 2, 1)
    Call .Move(m_ItemTextLeft + m_ItemOffset, (m_ListIndex - Bar) * m_tmpItemHeight + m_ItemOffset, m_ItemRct(m_ListIndex - Bar).x2 - m_ItemTextLeft - 2 * m_ItemOffset, m_tmpItemHeight - 2 * m_ItemOffset)
   Else
    .Alignment = 0
    Call .Move(m_ItemTextLeft + m_ItemOffset, (m_ListIndex - Bar) * m_tmpItemHeight + 0.5 * (m_tmpItemHeight - iScr.TextHeight("")), m_ItemRct(m_ListIndex - Bar).x2 - m_ItemTextLeft - 2 * m_ItemOffset, 1)
   End If
   '-- Get item text and turn TextBox to visible
   .Text = m_List(m_ListIndex).Text
   .SelStart = 0
   .SelLength = Len(txtEdit)
   .Visible = -1
   .SetFocus
  End With
 End If
End Sub

Public Sub EndEdit(Optional ByVal Modify As Boolean = 0)
 If (Modify = True) Then Call txtEdit_KeyPress(13) Else Call txtEdit_LostFocus
End Sub

Private Function APIRectangle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional ByVal lColor As OLE_COLOR = -1) As Long
 Dim hPen As Long, hPenOld As Long
 Dim PT   As POINTAPI
 
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(hDC, hPen)
 Call MoveToEx(hDC, X, Y, PT)
 Call LineTo(hDC, X + W, Y)
 Call LineTo(hDC, X + W, Y + H)
 Call LineTo(hDC, X, Y + H)
 Call LineTo(hDC, X, Y)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Function

Public Property Get SelectBorderColor() As OLE_COLOR
 SelectBorderColor = m_SelectBorderColor
End Property

Public Property Let SelectBorderColor(ByVal New_Color As OLE_COLOR)
 m_SelectBorderColor = GetLngColor(New_Color)
 Call PropertyChanged("SelectBorderColor")
End Property

Public Property Get SelectListBorderColor() As OLE_COLOR
 SelectListBorderColor = m_SelectListBorderColor
End Property

Public Property Let SelectListBorderColor(ByVal New_Color As OLE_COLOR)
 m_SelectListBorderColor = GetLngColor(New_Color)
 Call PropertyChanged("SelectListBorderColor")
End Property

Public Property Get ListGradient() As Boolean
 ListGradient = m_ListGradient
End Property

Public Property Let ListGradient(ByVal New_Gradient As Boolean)
 m_ListGradient = New_Gradient
 Call PropertyChanged("ListGradient")
End Property

Private Sub DrawGrad()
 Dim tmpRect As RECT2
 
 If (m_ListGradient = True) Then
  tmpRect.x1 = 0
  tmpRect.y1 = 0
  tmpRect.x2 = iScr.ScaleWidth
  tmpRect.y2 = iScr.ScaleHeight
  Call DrawBackGrad(iScr.hDC, tmpRect, m_ColorGradient1, m_ColorGradient2, GRADIENT_FILL_RECT_V)
 End If
End Sub

Public Sub Refresh()
 Call ReadjustBar
 Call DrawList
End Sub

Public Property Get ShadowColorText() As OLE_COLOR
 ShadowColorText = m_ShadowColorText
End Property

Public Property Let ShadowColorText(ByVal New_Color As OLE_COLOR)
 m_ShadowColorText = GetLngColor(New_Color)
 Call PropertyChanged("ShadowColorText")
End Property

Public Property Get VisibleRows() As Long
 VisibleRows = m_VisibleRows
End Property

Public Property Let VisibleRows(ByVal ItemsRows As Long)
 m_VisibleRows = ItemsRows
 Call PropertyChanged("VisibleRows")
End Property

'*------------------------------------------------------------------------------------>
'* Formatted using this program FormatCode 1.0.
'* Download a free version in this web www.geocities.com/hackprotm/FormatCode1.0.zip
'*------------------------------------------------------------------------------------>
