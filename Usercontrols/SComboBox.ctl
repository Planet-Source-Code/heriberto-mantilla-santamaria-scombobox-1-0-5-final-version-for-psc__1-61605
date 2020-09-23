VERSION 5.00
Begin VB.UserControl SComboBox 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   KeyPreview      =   -1  'True
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   147
   ToolboxBitmap   =   "SComboBox.ctx":0000
   Begin ComboBox.CoolList picList 
      Height          =   900
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   1588
      Appearance      =   0
      BorderStyle     =   0
      ScrollBarWidth  =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverSelection  =   -1  'True
      WordWrap        =   0   'False
      ItemHeight      =   20
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      ItemTextLeft    =   0
      ShadowColorText =   6582129
      VisibleRows     =   3
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   -1800
      TabIndex        =   0
      Top             =   270
      Width           =   255
   End
End
Attribute VB_Name = "SComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
'*                   Version 1.0.5                     *'
'*******************************************************'
'* Control:       SComboBox                            *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Collaboration: fred.cpp                             *'
'*                                                     *'
'*                So many thanks for his contribution  *'
'*                for this project, some styles and    *'
'*                Traduction to English of some        *'
'*                comments.                            *'
'*-----------------------------------------------------*'
'*                Credits and Thanks to                *'
'*-----------------------------------------------------*'
'*                Again my sincere gratefulness to     *'
'*                Paul Caton for it's spectacular      *'
'*                self-subclassing usercontrol         *'
'*                template, please see the             *'
'*                [CodeId = 54117].                    *'
'*-----------------------------------------------------*'
'*                MArio Florez for his Icon_Gray pictu-*'
'*                re routine please see the            *'
'*                [CodeId = 58622].                    *'
'*-----------------------------------------------------*'
'*                Carles P.V. for his excelent control *'
'*                CoolList OCX 1.2 please see the      *'
'*                [CodeId = 29586].                    *'
'*-----------------------------------------------------*'
'*            For suggestions and debugging            *'
'*-----------------------------------------------------*'
'* Dennis, Luciano, Amir and PSC community.            *'
'*******************************************************'
'* Description:   This usercontrol simulates a Combo-  *'
'*                Box But adds new an great features   *'
'*                like:                                *'
'*                                                     *'
'*                - When the list is shown doesn't     *'
'*                  deactivate the parent form.        *'
'*                - More than 20 Visual Styles; no     *'
'*                  images Everything done by code.    *'
'*                - Some extra cool properties.        *'
'*******************************************************'
'* Started on:    Friday, 11-jun-2004.                 *'
'*******************************************************'
'* Note:     Comments, suggestions, doubts or bug      *'
'*           reports are wellcome to these e-mail      *'
'*           addresses:                                *'
'*                                                     *'
'*                  heri_05-hms@mixmail.com or         *'
'*                  hcammus@hotmail.com                *'
'*                                                     *'
'*        Please rate my work on this control.         *'
'*    That lives the Soccer and the América of Cali    *'
'*             Of Colombia for the world.              *'
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
Option Explicit
 
'*******************************************************'
'* Subclasser Declarations Paul Caton                  *'
  
 '-uSelfSub declarations---------------------------------------------------------------------------
 Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
 End Enum

 Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
 Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
 Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
 Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
 Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
 Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
 Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
 Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
 Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
 Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

 Private z_ScMem             As Long                                         'Thunk base address
 Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
 Private z_Funk              As Collection                                   'hWnd/thunk-address collection
 
 Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
 Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
 Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
 Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
 Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
 Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
 Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
 Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
 Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 
 Public Event MouseEnter()
 Public Event MouseLeave()

 Private Const WM_ACTIVATE           As Long = &H6
 Private Const WM_MOUSEMOVE          As Long = &H200
 Private Const WM_MOUSELEAVE         As Long = &H2A3
 Private Const WM_MOVING             As Long = &H216
 Private Const WM_SIZING             As Long = &H214
 Private Const WM_EXITSIZEMOVE       As Long = &H232
 Private Const WM_LBUTTONDOWN        As Long = &H201
 Private Const WM_RBUTTONDOWN        As Long = &H204
 Private Const WM_MBUTTONDOWN        As Long = &H207
 Private Const WM_NCLBUTTONDOWN      As Long = &HA1
 Private Const WM_MOUSEWHEEL         As Long = &H20A
 Private Const WM_THEMECHANGED      As Long = &H31A

 Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
 End Enum

 Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                      As Long
  dwFlags                     As TRACKMOUSEEVENT_FLAGS
  hwndTrack                   As Long
  dwHoverTime                 As Long
 End Type

 Private bTrack                As Boolean
 Private bTrackUser32          As Boolean
 Private bInCtrl               As Boolean
 
 Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
 Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
 Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
'*******************************************************'
  
 '****************************'
 '* English: Private Type.   *'
 '* Español: Tipos Privados. *'
 '****************************'
 Private Type GRADIENT_RECT
  UpperLeft   As Long
  LowerRight  As Long
 End Type
   
 Private Type POINTAPI
  X           As Long
  Y           As Long
 End Type
 
 Private Type RECT
  Left       As Long
  Top        As Long
  Right      As Long
  Bottom     As Long
 End Type
 
 Private Type RGB
  Red         As Integer
  Green       As Integer
  Blue        As Integer
 End Type
 
 '* English: Elements of the list.
 '* Español: Elementos de la lista.
 Private Type PropertyCombo
  Color         As OLE_COLOR   '* Color of Text.
  Enabled       As Boolean     '* Item Enabled or Disabled.
  Image         As StdPicture  '* Item image.
  ImagePos      As Long        '* Position of Picture in the ImageList.
  Index         As Long        '* Index item.
  MouseIcon     As StdPicture  '* Set MouseIcon for each item.
  SeparatorLine As Boolean     '* Set SeparatorLine for each group that you consider necessary.
  Tag           As String      '* Extra Information only if is necessary.
  Text          As String      '* Text of the item.
  TextShadow    As Boolean     '* Shadow text item.
  ToolTipText   As String      '* ToolTipText for item.
 End Type
  
 Private Type TRIVERTEX
  X             As Long
  Y             As Long
  Red           As Integer
  Green         As Integer
  Blue          As Integer
  Alpha         As Integer
 End Type
  
 '*********************************************'
 '* English: Public Enum of Control.          *'
 '* Español: Enumeración Publica del control. *'
 '*********************************************'
 
 '* English: Enum for the alignment of the text of the list.
 '* Español: Enum para la alineación del texto de la lista.
 Public Enum AlignTextCombo
  AlignLeft = &H0
  AlignRight = &H1
  AlignCenter = &H2
 End Enum
 
 '* English: Appearance Combo.
 '* Español: Apariencias del Combo.
 Public Enum ComboAppearance
  Office = &H1             '* By fred.cpp & HACKPRO TM.
  Win98 = &H2              '* By fred.cpp.
  WinXp = &H3              '* By fred.cpp & HACKPRO TM.
  [Soft Style] = &H4       '* By fred.cpp.
  KDE = &H5                '* By HACKPRO TM.
  Mac = &H6                '* By fred.cpp & HACKPRO TM.
  JAVA = &H7               '* By fred.cpp.
  [Explorer Bar] = &H8     '* By HACKPRO TM.
  Picture = &H9            '* By HACKPRO TM.
  [Special Borde] = &HA    '* By HACKPRO TM.
  Circular = &HB           '* By HACKPRO TM.
  [GradientV] = &HC        '* By HACKPRO TM.
  [GradientH] = &HD        '* By HACKPRO TM.
  [Light Blue] = &HE       '* By HACKPRO TM.
  [Style Arrow] = &HF      '* By HACKPRO TM.
  [NiaWBSS] = &H10         '* By HACKPRO TM.
  [Rhombus] = &H11         '* By HACKPRO TM.
  [Additional Xp] = &H12   '* By HACKPRO TM.
  [Ardent] = &H13          '* By HACKPRO TM.
  [Chocolate] = &H14       '* By HACKPRO TM.
  [Button Download] = &H15 '* By HACKPRO TM.
 End Enum

 '* English: Type of Combo and behavior of the list.
 '* Español: Tipo de Combo y comportamiento de la lista.
 Public Enum ComboStyle
  [Dropdown Combo] = &H0
  [Dropdown List] = &H1
 End Enum
 
 '* English: Appearance standard style Office.
 '* Español: Apariencias estándares del estilo Office.
 Public Enum ComboOfficeAppearance
  [Office Xp] = &H0       '* By HACKPRO TM.
  [Office 2000] = &H1     '* By fred.cpp.
  [Office 2003] = &H2     '* By HACKPRO TM.
 End Enum
 
 '* English: Appearance standard style Xp.
 '* Español: Apariencias estándares del estilo Xp.
 Public Enum ComboXpAppearance
  [Windows Themed] = &H0  '* By fred.cpp
  Aqua = &H1              '* By HACKPRO TM.
  [Olive Green] = &H2     '* By HACKPRO TM.
  Silver = &H3            '* By HACKPRO TM.
  TasBlue = &H4           '* By HACKPRO TM.
  Gold = &H5              '* By HACKPRO TM.
  Blue = &H6              '* By HACKPRO TM.
  CustomXP = &H7          '* By HACKPRO TM.
 End Enum
  
 '* English: Direction of like the list is shown.
 '* Español: Dirección de como se muestra la lista.
 Public Enum ListDirection
  [Show Down] = &H0
  [Show Up] = &H1
 End Enum
  
 '* English: Enum for the type of text comparison.
 '* Español: Enum para el tipo de comparación de texto.
 Public Enum StringCompare
  NoneWord = &H0
  ExactWord = &H1
  CompleteWord = &H2
 End Enum
  
 '********************************'
 '* English: Private variables.  *'
 '* Español: Variables privadas. *'
 '********************************'
 Private BigText                 As String
 Private ControlEnabled          As Boolean
 Private cValor                  As Long
 Private g_Font                  As StdFont
 Private iFor                    As Long
 Private Images                  As Object
 Private IndexItemNow            As Long
 Private isFailedXP              As Boolean
 Private isPicture               As Boolean
 Private isScroll                As Boolean
 Private ItemFocus               As Long
 Private KeyPos                  As Integer
 Private ListContents()          As PropertyCombo
 Private ListMaxL                As Long
 Private m_btnRect               As RECT
 Private m_StateG                As Integer
 Private m_TextRct()             As RECT
 Private myAlignCombo            As AlignTextCombo
 Private myAppearanceCombo       As ComboAppearance
 Private myArrowColor            As OLE_COLOR
 Private myAutoSel               As Boolean
 Private myBackColor             As OLE_COLOR
 Private myDisabledColor         As OLE_COLOR
 Private myDisabledPictureUser   As StdPicture
 Private myFocusPictureUser      As StdPicture
 Private myGradientColor1        As OLE_COLOR
 Private myGradientColor2        As OLE_COLOR
 Private myHighLightBorderColor  As OLE_COLOR
 Private myHighLightColorText    As OLE_COLOR
 Private myHighLightPictureUser  As StdPicture
 Private myItemsShow             As Long
 Private myListColor             As OLE_COLOR
 Private myListGradient          As Boolean
 Private myListShown             As ListDirection
 Private myMouseIcon             As StdPicture
 Private myMousePointer          As MousePointerConstants
 Private myNormalBorderColor     As OLE_COLOR
 Private myNormalColorText       As OLE_COLOR
 Private myNormalPictureUser     As StdPicture
 Private myOfficeAppearance      As ComboOfficeAppearance
 Private mySelectBorderColor     As OLE_COLOR
 Private mySelectListBorderColor As OLE_COLOR
 Private mySelectListColor       As OLE_COLOR
 Private myShadowColorText       As OLE_COLOR
 Private myStyleCombo            As ComboStyle
 Private myText                  As String
 Private myXpAppearance          As ComboXpAppearance
 Private NoShow                  As Boolean
 Private OrderListContents()     As PropertyCombo
 Private RGBColor                As RGB
 Private sumItem                 As Long
 Private tempBorderColor         As OLE_COLOR
 Private ThisFocus               As Long
 Private tmpC1                   As Long
 Private tmpC2                   As Long
 Private tmpC3                   As Long
 Private tmpColor                As Long
 Private UserText                As String
  
 '***************************************'
 '* English: Constant declares.         *'
 '* Español: Declaración de Constantes. *'
 '***************************************'
 Private Const BDR_RAISEDINNER = &H4
 Private Const BDR_SUNKENOUTER = &H2
 Private Const BF_RECT = (&H1 Or &H2 Or &H4 Or &H8)
 Private Const COLOR_BTNFACE = 15
 Private Const COLOR_BTNSHADOW = 16
 Private Const COLOR_GRADIENTACTIVECAPTION As Long = 27
 Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
 Private Const COLOR_GRAYTEXT As Long = 17
 Private Const COLOR_HOTLIGHT As Long = 26
 Private Const COLOR_INACTIVECAPTIONTEXT As Long = 19
 Private Const COLOR_WINDOW = 5
 Private Const defAppearanceCombo = 1
 Private Const defArrowColor = &HC56A31
 Private Const defDisabledColor = &H808080
 Private Const defGradientColor1 = &HDAB278
 Private Const defGradientColor2 = &HFFDD9E
 Private Const defHighLightBorderColor = &HC56A31
 Private Const defHighLightColorText = &HFFFFFF
 Private Const defNormalBorderColor = &HDEEDEF
 Private Const defNormalColorText = &HC56A31
 Private Const defListColor = &HFFFFFF
 Private Const defListShown = 0
 Private Const defOfficeAppearance = 0
 Private Const defSelectBorderColor = &HC56A31
 Private Const defSelectListBorderColor = &H6B2408
 Private Const defSelectListColor = &HC56A31
 Private Const defShadowColorText = &H80000015
 Private Const defStyleCombo = 0
 Private Const DSS_DISABLED = &H20
 Private Const DST_BITMAP = &H3
 Private Const DST_COMPLEX = &H0
 Private Const DST_ICON = &H3
 Private Const DT_LEFT                As Long = &H0
 Private Const DT_SINGLELINE          As Long = &H20
 Private Const DT_VCENTER             As Long = &H4
 Private Const DT_WORD_ELLIPSIS       As Long = &H40000
 Private Const EDGE_RAISED = (&H1 Or &H4)
 Private Const EDGE_SUNKEN = (&H2 Or &H8)
 Private Const GRADIENT_FILL_RECT_H   As Long = &H0
 Private Const GRADIENT_FILL_RECT_V   As Long = &H1
 Private Const GWL_EXSTYLE = -20
 Private Const SWP_FRAMECHANGED = &H20
 Private Const SWP_NOMOVE = &H2
 Private Const SWP_NOSIZE = &H1
 Private Const WS_EX_TOOLWINDOW = &H80
 Private Const Version                As String = "SComboBox 1.0.5 By HACKPRO TM"
 
 '*******************************'
 '* English: Private WithEvents *'
 '* Español: Private WithEvents *'
 '*******************************'
 Private WithEvents picTemp       As PictureBox
Attribute picTemp.VB_VarHelpID = -1
 
 '******************************'
 '* English: Public Events.    *'
 '* Español: Eventos Públicos. *'
 '******************************'
 Public Event SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
 Public Event TotalItems(ByVal ListCount As Long)
 
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
    
 '**********************************'
 '* English: Calls to the API's.   *'
 '* Español: Llamadas a los API's. *'
 '**********************************'
 Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
 Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
 Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
 Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
 Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
 Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
 Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
 Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
 Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
 Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
 Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
 Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
  
Private Sub picList_Click()
 If (sumItem > 0) Then
  With picList
   ItemFocus = .ListIndex + 1
   If (ListContents(ListIndex).Enabled = True) Then
    .Visible = False
    Text = ListContents(ListIndex).Text
    NoShow = False
    bInCtrl = False
    RaiseEvent SelectionMade(ListContents(ListIndex).Text, ListIndex)
   End If
  End With
 End If
End Sub

Private Sub txtCombo_Change()
 Dim sItem As Long, iLen As Long, iStart As Long
 
On Error Resume Next
 iStart = txtCombo.SelStart
 If (myAutoSel = False) Then
  sItem = FindItemText(txtCombo.Text, 2)
  If (sItem > 0) Then
   If (ListContents(sItem).Enabled = True) Then
    ItemFocus = sItem
    IndexItemNow = sItem
    If (IndexItemNow > NumberItemsToShow) Then
     iLen = (NumberItemsToShow + IndexItemNow) - IndexItemNow
    Else
     iLen = IndexItemNow - (NumberItemsToShow + IndexItemNow)
    End If
   End If
  Else
   ItemFocus = -1
   picList.ListIndex = -1
  End If
 ElseIf (KeyPos <> 0) And (KeyPos <> 67) And (KeyPos <> 46) Then
  sItem = FindItemText(txtCombo.Text)
  If (sItem > 0) Then
   iLen = Len(txtCombo.Text)
   txtCombo.Text = txtCombo.Text & Mid$(ListContents(sItem).Text, iLen + 1, Len(ListContents(sItem).Text))
   txtCombo.SelStart = iLen
   txtCombo.SelLength = Len(txtCombo.Text)
   sItem = FindItemText(txtCombo.Text, 2)
   If (sItem > 0) Then
    If (ListContents(sItem).Enabled = True) Then
     ItemFocus = sItem
     IndexItemNow = sItem
    End If
   Else
    ItemFocus = -1
    picList.ListIndex = -1
   End If
  Else
   ItemFocus = -1
  End If
 Else
  ItemFocus = FindItemText(txtCombo.Text, 2)
  txtCombo.SelStart = iStart
 End If
 Call isEnabled(ControlEnabled)
End Sub

Private Sub txtCombo_Click()
 If (picList.Visible = True) Then
 On Error Resume Next
  picList.Visible = False
  txtCombo.SetFocus
 End If
End Sub

Private Sub txtCombo_GotFocus()
 txtCombo.SelStart = 0
 txtCombo.SelLength = Len(txtCombo.Text)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
 If (AppearanceCombo = 18) Then Call isEnabled(ControlEnabled)
End Sub

Private Sub UserControl_ExitFocus()
 Call isEnabled(ControlEnabled)
 bInCtrl = False
 NoShow = False
End Sub

Private Sub UserControl_Initialize()
 Dim OS As OSVERSIONINFO

 '* Get the operating system version for text drawing purposes.
 OS.dwOSVersionInfoSize = Len(OS)
 Call GetVersionEx(OS)
 mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
End Sub

Private Sub UserControl_InitProperties()
 '* English: Setup properties values.
 '* Español: Establece propiedades iniciales.
 ControlEnabled = True
 ItemFocus = -1
 isPicture = False
 ListIndex = -1
 ListMaxL = 10
 myListShown = 0
 myAutoSel = False
 myAppearanceCombo = defAppearanceCombo
 myArrowColor = defArrowColor
 myBackColor = defListColor
 myDisabledColor = defDisabledColor
 myGradientColor1 = defGradientColor1
 myGradientColor2 = defGradientColor2
 myHighLightBorderColor = defHighLightBorderColor
 myHighLightColorText = defHighLightColorText
 myItemsShow = 7
 myListColor = defListColor
 myListGradient = False
 myNormalBorderColor = defNormalBorderColor
 myNormalColorText = defNormalColorText
 myOfficeAppearance = defOfficeAppearance
 mySelectBorderColor = defSelectBorderColor
 mySelectListBorderColor = defSelectListBorderColor
 mySelectListColor = defSelectListColor
 myShadowColorText = defShadowColorText
 myStyleCombo = defStyleCombo
 myText = Ambient.DisplayName
 Text = myText
 myXpAppearance = 1
 Set g_Font = Ambient.Font
 sumItem = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim LastIndex As Integer
 
 LastIndex = ListIndex - 1
 Select Case KeyCode
  Case 38 '{Up arrow}
   If (LastIndex > 1) Then LastIndex = LastIndex - 1
  Case 40 '{Down arrow}
   If (LastIndex < ListCount - 1) Then LastIndex = LastIndex + 1
  Case 33 '{PageDown}
   If (LastIndex > NumberItemsToShow) Then
    LastIndex = LastIndex - NumberItemsToShow
   Else
    LastIndex = 1
   End If
  Case 34 '{PageUp}
   If (LastIndex < ListCount - NumberItemsToShow - 1) Then
    LastIndex = LastIndex + NumberItemsToShow
   Else
    LastIndex = ListCount - 1
   End If
  Case 36 '{Start}
   LastIndex = 1
  Case 35 '{End}
   LastIndex = ListCount - 1
  Case 13
   Call picList_Click
 End Select
 ListIndex = LastIndex
End Sub

Private Sub UserControl_LostFocus()
 Call UserControl_ExitFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim oRect As RECT
 
 '* English: Show or hide the list.
 '* Español: Muestra la lista ó la oculta.
 If (Button = vbLeftButton) And (picList.Visible = False) Then
  IndexItemNow = ListIndex - 1
  Call DrawAppearance(myAppearanceCombo, 3)
  If (myAppearanceCombo = 2) Or ((myAppearanceCombo = 3) And ((myXpAppearance = 7) Or (myXpAppearance = 0))) Then
   Call Espera(0.09)
   Call DrawAppearance(myAppearanceCombo, 1)
  End If
  If (txtCombo.Visible = True) Then
   ItemFocus = picList.FindFirst(txtCombo.Text, , True) + 1
  Else
   ItemFocus = picList.FindFirst(UserText, , True) + 1
  End If
  If (ItemFocus = -1) Then
   Call picList.AddItem("", , , , False)
   ItemFocus = 0
  End If
  If (txtCombo.Visible = True) And (txtCombo.Text = "") Then ItemFocus = 0
  ThisFocus = ItemFocus
  Call GetWindowRect(UserControl.hWnd, oRect)
  picList.Width = ScaleWidth
  If (myListShown = 1) Then
   '* The list is shown up.
   Call picList.Move(oRect.Left, (oRect.Bottom - (picList.Height + UserControl.Height + 1)))
  Else
   '* The list is shown down.
   Call picList.Move(oRect.Left, oRect.Bottom + 1)
  End If
  If (NumberItemsToShow < MaxListLength) Then
   isScroll = True
  Else
   isScroll = False
  End If
  Call SetWindowPos(picList.hWnd, -1, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
  NoShow = True
  Call DrawList(ItemFocus - 1)
  picList.Visible = True
 Else
  picList.Visible = False
  NoShow = False
  bInCtrl = False
  Call DrawAppearance(myAppearanceCombo, 2)
 End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Alignment = PropBag.ReadProperty("Alignment", 0)
 AppearanceCombo = PropBag.ReadProperty("AppearanceCombo", defAppearanceCombo)
 ArrowColor = PropBag.ReadProperty("ArrowColor", defArrowColor)
 AutoCompleteWord = PropBag.ReadProperty("AutoCompleteWord", False)
 BackColor = PropBag.ReadProperty("BackColor", defListColor)
 Call ControlsSubClasing
 DisabledColor = PropBag.ReadProperty("DisabledColor", defDisabledColor)
 Set DisabledPictureUser = PropBag.ReadProperty("DisabledPictureUser", Nothing)
 Enabled = PropBag.ReadProperty("Enabled", True)
 GradientColor1 = PropBag.ReadProperty("GradientColor1", defGradientColor1)
 GradientColor2 = PropBag.ReadProperty("GradientColor2", defGradientColor2)
 Set FocusPictureUser = PropBag.ReadProperty("FocusPictureUser", Nothing)
 Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
 HighLightBorderColor = PropBag.ReadProperty("HighLightBorderColor", defHighLightBorderColor)
 HighLightColorText = PropBag.ReadProperty("HighLightColorText", defHighLightColorText)
 Set HighLightPictureUser = PropBag.ReadProperty("HighLightPictureUser", Nothing)
 ListColor = PropBag.ReadProperty("ListColor", defListColor)
 ListGradient = PropBag.ReadProperty("ListGradient", False)
 ListPositionShow = PropBag.ReadProperty("ListPositionShow", defListShown)
 MaxListLength = PropBag.ReadProperty("MaxListLength", "10")
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 MousePointer = PropBag.ReadProperty("MousePointer", 0)
 NormalBorderColor = PropBag.ReadProperty("NormalBorderColor", defNormalBorderColor)
 NormalColorText = PropBag.ReadProperty("NormalColorText", defNormalColorText)
 Set NormalPictureUser = PropBag.ReadProperty("NormalPictureUser", Nothing)
 NumberItemsToShow = PropBag.ReadProperty("NumberItemsToShow", "7")
 OfficeAppearance = PropBag.ReadProperty("OfficeAppearance", defOfficeAppearance)
 SelectBorderColor = PropBag.ReadProperty("SelectBorderColor", defSelectBorderColor)
 SelectListBorderColor = PropBag.ReadProperty("SelectListBorderColor", defSelectListBorderColor)
 SelectListColor = PropBag.ReadProperty("SelectListColor", defSelectListColor)
 ShadowColorText = PropBag.ReadProperty("ShadowColorText", defShadowColorText)
 Style = PropBag.ReadProperty("Style", defStyleCombo)
 Text = PropBag.ReadProperty("Text", Ambient.DisplayName)
 XpAppearance = PropBag.ReadProperty("XpAppearance", 1)
 If (Ambient.UserMode = True) Then
  bTrack = True
  bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  If Not (bTrackUser32 = True) Then
   If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
    bTrack = False
   End If
  End If
  If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
   '* Start subclassing the UserControl.
   With UserControl
    Call sc_Subclass(.hWnd)
    Call sc_AddMsg(.hWnd, WM_MOUSEMOVE)
    Call sc_AddMsg(.hWnd, WM_MOUSELEAVE)
    Call sc_AddMsg(.hWnd, WM_MOUSEWHEEL)
    Call sc_Subclass(txtCombo.hWnd)
    Call sc_AddMsg(txtCombo.hWnd, WM_MOUSEMOVE)
    Call sc_AddMsg(txtCombo.hWnd, WM_MOUSELEAVE)
   End With
   With UserControl.Parent
    Call sc_Subclass(.hWnd)
    Call sc_AddMsg(.hWnd, WM_MOUSEMOVE)
    Call sc_AddMsg(.hWnd, WM_MOUSELEAVE)
    Call sc_AddMsg(.hWnd, WM_MOVING)
    Call sc_AddMsg(.hWnd, WM_SIZING)
    Call sc_AddMsg(.hWnd, WM_EXITSIZEMOVE)
    Call sc_AddMsg(.hWnd, WM_LBUTTONDOWN)
    Call sc_AddMsg(.hWnd, WM_RBUTTONDOWN)
    Call sc_AddMsg(.hWnd, WM_MBUTTONDOWN)
    Call sc_AddMsg(.hWnd, WM_ACTIVATE)
    Call sc_AddMsg(.hWnd, WM_NCLBUTTONDOWN)
   End With
  End If
 End If
On Error GoTo 0
End Sub

Private Sub UserControl_Resize()
 If Not (Ambient.UserMode = True) Then Call isEnabled(ControlEnabled)
 Call isEnabled(ControlEnabled)
End Sub

Private Sub UserControl_Show()
 Dim lResult As Long
 
On Error Resume Next
 lResult = GetWindowLong(picList.hWnd, GWL_EXSTYLE)
 Call SetWindowLong(picList.hWnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
 Call SetWindowPos(picList.hWnd, picList.hWnd, 0, 0, 0, 0, 39)
 Call SetWindowLong(picList.hWnd, -8, Parent.hWnd)
 Call SetParent(picList.hWnd, 0)
 If (isPicture = False) Then txtCombo.Left = 8
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch
 Erase ListContents
 Set picTemp = Nothing
 Call sc_Terminate '* Stop all subclassing.
 Exit Sub
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Alignment", myAlignCombo, 0)
 Call PropBag.WriteProperty("AppearanceCombo", myAppearanceCombo, defAppearanceCombo)
 Call PropBag.WriteProperty("ArrowColor", myArrowColor, defArrowColor)
 Call PropBag.WriteProperty("AutoCompleteWord", myAutoSel, False)
 Call PropBag.WriteProperty("BackColor", myBackColor, defListColor)
 Call PropBag.WriteProperty("DisabledColor", myDisabledColor, defDisabledColor)
 Call PropBag.WriteProperty("DisabledPictureUser", myDisabledPictureUser, Nothing)
 Call PropBag.WriteProperty("Enabled", ControlEnabled, True)
 Call PropBag.WriteProperty("FocusPictureUser", myFocusPictureUser, Nothing)
 Call PropBag.WriteProperty("Font", g_Font, Ambient.Font)
 Call PropBag.WriteProperty("GradientColor1", myGradientColor1, defGradientColor1)
 Call PropBag.WriteProperty("GradientColor2", myGradientColor2, defGradientColor2)
 Call PropBag.WriteProperty("HighLightBorderColor", myHighLightBorderColor, defHighLightBorderColor)
 Call PropBag.WriteProperty("HighLightColorText", myHighLightColorText, defHighLightColorText)
 Call PropBag.WriteProperty("HighLightPictureUser", myHighLightPictureUser, Nothing)
 Call PropBag.WriteProperty("ListColor", myListColor, defListColor)
 Call PropBag.WriteProperty("ListGradient", myListGradient, False)
 Call PropBag.WriteProperty("ListPositionShow", myListShown, defListShown)
 Call PropBag.WriteProperty("MaxListLength", ListMaxL, "10")
 Call PropBag.WriteProperty("MouseIcon", myMouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", myMousePointer, 0)
 Call PropBag.WriteProperty("NormalBorderColor", myNormalBorderColor, defNormalBorderColor)
 Call PropBag.WriteProperty("NormalColorText", myNormalColorText, defNormalColorText)
 Call PropBag.WriteProperty("NormalPictureUser", myNormalPictureUser, Nothing)
 Call PropBag.WriteProperty("NumberItemsToShow", myItemsShow, "7")
 Call PropBag.WriteProperty("OfficeAppearance", myOfficeAppearance, defOfficeAppearance)
 Call PropBag.WriteProperty("SelectBorderColor", mySelectBorderColor, defSelectBorderColor)
 Call PropBag.WriteProperty("SelectListBorderColor", mySelectListBorderColor, defSelectListBorderColor)
 Call PropBag.WriteProperty("SelectListColor", mySelectListColor, defSelectListColor)
 Call PropBag.WriteProperty("ShadowColorText", myShadowColorText, defShadowColorText)
 Call PropBag.WriteProperty("Style", myStyleCombo, defStyleCombo)
 Call PropBag.WriteProperty("Text", myText, Ambient.DisplayName)
 Call PropBag.WriteProperty("XpAppearance", myXpAppearance, 1)
End Sub

'*******************************************'
'* English: Properties of the Usercontrol. *'
'* Español: Propiedades del Usercontrol.   *'
'*******************************************'
Public Property Get Alignment() As AlignTextCombo
 '* English: Sets/Gets alignment of the text in the list.
 '* Español: Devuelve o establece la alineación del texto en la lista.
 Alignment = myAlignCombo
End Property

Public Property Let Alignment(ByVal New_Align As AlignTextCombo)
 Dim isAlign As Integer
 
 myAlignCombo = New_Align
 If (New_Align = AlignCenter) Then
  isAlign = 1
 ElseIf (New_Align = AlignLeft) Then
  isAlign = 0
 Else
  isAlign = 2
 End If
 picList.Alignment = isAlign
 Call PropertyChanged("Alignment")
 Refresh
End Property

Public Property Get AppearanceCombo() As ComboAppearance
 '* English: Sets/Gets the style of the Combo.
 '* Español: Devuelve o establece el estilo del Combo.
 AppearanceCombo = myAppearanceCombo
End Property

Public Property Let AppearanceCombo(ByVal New_Style As ComboAppearance)
 myAppearanceCombo = IIf(New_Style <= 0, 1, New_Style)
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("AppearanceCombo")
 Refresh
End Property

Public Property Get ArrowColor() As OLE_COLOR
 '* English: Sets/Gets the color of the arrow.
 '* Español: Devuelve o establece el color de la flecha.
 ArrowColor = myArrowColor
End Property

Public Property Let ArrowColor(ByVal New_Color As OLE_COLOR)
 myArrowColor = ConvertSystemColor(New_Color)
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("ArrowColor")
 Refresh
End Property

Public Property Get AutoCompleteWord() As Boolean
 '* English: Sets/Gets complete the word with a similar element of the list.
 '* Español: Devuelve o establece si se completa la palabra con un elemento similar de la lista.
 AutoCompleteWord = myAutoSel
End Property
'* Note: When this property this active one and the list _
         is shown, it is not tried to locate the element _
         in the list to make quicker the search of the _
         text to complete.
'* Nota: Cuando esta propiedad este activa y la lista se _
         muestre, no se intentara ubicar el elemento en la _
         lista para hacer más rápido la búsqueda del texto _
         a completar.

Public Property Let AutoCompleteWord(ByVal New_Value As Boolean)
 myAutoSel = New_Value
 Call PropertyChanged("AutoCompleteWord")
 Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
 '* English: Sets/Gets the color of the Usercontrol.
 '* Español: Devuelve o establece el color del Usercontrol.
 BackColor = myBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
 myBackColor = ConvertSystemColor(GetLngColor(New_Color))
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("BackColor")
 Refresh
End Property

Public Property Get DisabledColor() As OLE_COLOR
 '* English: Sets/Gets the color of the disabled text.
 '* Español: Devuelve o establece el color del texto deshabilitado.
 DisabledColor = ShiftColorOXP(myDisabledColor, 94)
End Property

Public Property Let DisabledColor(ByVal New_Color As OLE_COLOR)
 myDisabledColor = ConvertSystemColor(GetLngColor(New_Color))
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("DisabledColor")
 Refresh
End Property

Public Property Get DisabledPictureUser() As StdPicture
 '* English: Sets/Gets an image like topic of the Combo when the Object is not enabled.
 '* Español: Devuelve o establece una imagen como tema del combo cuando el Objeto este inactivo.
 Set DisabledPictureUser = myDisabledPictureUser
End Property

Public Property Set DisabledPictureUser(ByVal New_Picture As StdPicture)
 Set myDisabledPictureUser = New_Picture
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("DisabledPictureUser")
 Refresh
End Property

Public Property Get Enabled() As Boolean
 '* English: Sets/Gets the Enabled property of the control.
 '* Español: Devuelve o establece si el Usercontrol esta habilitado ó deshabilitado.
 Enabled = ControlEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 UserControl.Enabled = New_Enabled
 ControlEnabled = New_Enabled
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("Enabled")
 Refresh
End Property

Public Property Get FocusPictureUser() As StdPicture
 '* English: Sets/Gets the image like topic of the Combo when It has the focus.
 '* Español: Devuelve o establece una imagen como tema del combo cuando se tiene el enfoque.
 Set FocusPictureUser = myFocusPictureUser
End Property

Public Property Set FocusPictureUser(ByVal New_Picture As StdPicture)
 Set myFocusPictureUser = New_Picture
 Call PropertyChanged("FocusPictureUser")
 Refresh
End Property

Public Property Get Font() As StdFont
 '* English: Sets/Gets the Font of the control.
 '* Español: Devuelve o establece el tipo de fuente del texto.
 Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
On Error Resume Next
 With g_Font
  .Name = New_Font.Name
  .Size = IIf(New_Font.Size > 12, 8, New_Font.Size)
  .Bold = New_Font.Bold
  .Italic = New_Font.Italic
  .Underline = New_Font.Underline
  .Strikethrough = New_Font.Strikethrough
 End With
 txtCombo.Font = New_Font
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("Font")
 Refresh
End Property

Public Property Get GradientColor1() As OLE_COLOR
 '* English: Sets/Gets the color First gradient color.
 '* Español: Devuelve o establece el color Gradient 1.
 GradientColor1 = myGradientColor1
End Property

Public Property Let GradientColor1(ByVal New_Color As OLE_COLOR)
 myGradientColor1 = ConvertSystemColor(New_Color)
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("GradientColor1")
 Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
 '* English: Sets/Gets the Second gradient color.
 '* Español: Devuelve o establece el color Gradient 2.
 GradientColor2 = myGradientColor2
End Property

Public Property Let GradientColor2(ByVal New_Color As OLE_COLOR)
 myGradientColor2 = ConvertSystemColor(New_Color)
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("GradientColor2")
 Refresh
End Property

Public Property Get HighLightBorderColor() As OLE_COLOR
 '* English: Sets/Gets the color of the border of the control when the the control is highlighted.
 '* Español: Devuelve o establece el color del borde del control cuando el pasa sobre él.
 HighLightBorderColor = myHighLightBorderColor
End Property

Public Property Let HighLightBorderColor(ByVal New_Color As OLE_COLOR)
 myHighLightBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("HighLightBorderColor")
 Refresh
End Property

Public Property Get HighLightColorText() As OLE_COLOR
 '* English: Sets/Gets the color of the selection of the text.
 '* Español: Devuelve o establece el color de selección del texto.
 HighLightColorText = myHighLightColorText
End Property

Public Property Let HighLightColorText(ByVal New_Color As OLE_COLOR)
 myHighLightColorText = ConvertSystemColor(New_Color)
 Call PropertyChanged("HighLightColorText")
 Refresh
End Property

Public Property Get HighLightPictureUser() As StdPicture
 '* English: Sets/Gets an image like topic of the Combo when the mouse is over the control.
 '* Español: Devuelve o establece una imagen como tema del combo cuando el mouse pasa por el Objeto.
 Set HighLightPictureUser = myHighLightPictureUser
End Property

Public Property Set HighLightPictureUser(ByVal New_Picture As StdPicture)
 Set myHighLightPictureUser = New_Picture
 Call PropertyChanged("HighLightPictureUser")
 Refresh
End Property

Public Property Get hWnd() As Long
 '* English: Returns a handle to a form or control.
 '* Español: Devuelve el controlador de un formulario o un control.
 hWnd = UserControl.hWnd
End Property

Public Property Get ItemTag(ByVal ListIndex As Long) As String
 '* English: Returns the tag of a specified item.
 '* Español: Selecciona el tag de Item.
 ItemTag = ""
On Error GoTo myErr:
 ItemTag = ListContents(ListIndex).Tag
 Exit Property
myErr:
 ItemTag = ""
End Property

Public Property Get ListColor() As OLE_COLOR
 '* English: Sets/Gets the color of the List.
 '* Español: Devuelve o establece el color de la lista.
 ListColor = myListColor
End Property

Public Property Let ListColor(ByVal New_Color As OLE_COLOR)
 myListColor = ConvertSystemColor(New_Color)
 picList.BackNormal = myListColor
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("ListColor")
 Refresh
End Property

Public Property Get ListCount() As Long
 '* English: Returns the number of elements in the list.
 '* Español: Devuelve o establece el número de elementos de la lista.
 ListCount = picList.ListCount
End Property

Public Property Get ListGradient() As Boolean
 '* English: Sets/Gets the list in degraded form.
 '* Español: Devuelve o establece si la lista se muestra en forma degradada.
 ListGradient = myListGradient
End Property

Public Property Let ListGradient(ByVal New_Gradient As Boolean)
 myListGradient = New_Gradient
 Call PropertyChanged("ListGradient")
 Refresh
End Property

Public Property Get ListIndex() As Long
 '* English: Sets/Gets the selected item.
 '* Español: Devuelve o establece el item actual seleccionado.
 ListIndex = picList.ListIndex + 1
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Long)
On Error Resume Next
 picList.ListIndex = New_ListIndex
 ItemFocus = picList.ListIndex + 1
 Call List(ItemFocus)
 Text = ListContents(ItemFocus).Text
On Error GoTo 0
End Property

Public Property Get ListPositionShow() As ListDirection
 '* English: Sets/Gets If the list is shown up or down.
 '* Español: Devuelve o establece si la lista se muestra hacia arriba ó hacia abajo.
 ListPositionShow = myListShown
End Property

Public Property Let ListPositionShow(ByVal New_Position As ListDirection)
 myListShown = New_Position
 Call PropertyChanged("ListPositionShow")
 Refresh
End Property

Public Property Get MaxListLength() As Long
 '* English: Sets/Gets the maximum size of the list.
 '* Español: Devuelve o establece el tamaño máximo de la lista.
 MaxListLength = IIf(ListMaxL < 0, ListCount, ListMaxL)
End Property

Public Property Let MaxListLength(ByVal ListMax As Long)
 If (ListMax > 0) And (ListMax < picList.ListCount) Then
  ListMaxL = ListMax
 Else
  ListMaxL = ListCount
 End If
 Call PropertyChanged("MaxListLength")
 Refresh
End Property

Public Property Get MouseIcon() As StdPicture
 '* English: Sets a custom mouse icon.
 '* Español: Establece un icono escogido por el usuario.
 Set MouseIcon = myMouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
 Set myMouseIcon = New_MouseIcon
End Property

Public Property Get MousePointer() As MousePointerConstants
 '* English: Sets/Gets the type of mouse pointer displayed when over part of an object.
 '* Español: Devuelve o establece el tipo de puntero a mostrar cuando el mouse pase sobre el objeto.
 MousePointer = myMousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
 myMousePointer = New_MousePointer
End Property

Public Property Get NewIndex() As Long
 '* English: Sets/Gets the last Item added.
 '* Español: Devuelve o establece el último item agregado.
 If (sumItem <= 0) Then NewIndex = -1 Else NewIndex = ListContents(sumItem).Index
End Property

Public Property Get NormalBorderColor() As OLE_COLOR
 '* English: Sets/Gets the normal border color of the control.
 '* Español: Devuelve o establece el color normal del borde del control.
 NormalBorderColor = myNormalBorderColor
End Property

Public Property Let NormalBorderColor(ByVal New_Color As OLE_COLOR)
 myNormalBorderColor = ConvertSystemColor(New_Color)
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("NormalBorderColor")
 Refresh
End Property

Public Property Let NormalColorText(ByVal New_Color As OLE_COLOR)
 myNormalColorText = ConvertSystemColor(New_Color)
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("NormalColorText")
 Refresh
End Property

Public Property Get NormalColorText() As OLE_COLOR
 '* English: Sets/Gets the normal text color in the control.
 '* Español: Devuelve o establece el color del texto normal.
 NormalColorText = myNormalColorText
End Property

Public Property Get NormalPictureUser() As StdPicture
 '* English: Sets/Gets an image like topic of the Combo in normal state.
 '* Español: Devuelve o establece una imagen como tema del combo en estado normal.
 Set NormalPictureUser = myNormalPictureUser
End Property

Public Property Set NormalPictureUser(ByVal New_Picture As StdPicture)
 Set myNormalPictureUser = New_Picture
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("NormalPictureUser")
 Refresh
End Property

Public Property Get NumberItemsToShow() As Long
 '* English: Sets/Gets the number of items to show per time.
 '* Español: Devuelve o establece el número de items a mostrar por vez.
 If (myItemsShow < 0) Then myItemsShow = IIf(MaxListLength > 8, 7, MaxListLength)
 NumberItemsToShow = IIf(myItemsShow = 0, 8, myItemsShow)
End Property

Public Property Let NumberItemsToShow(ByVal ItemsShow As Long)
 If (ItemsShow <= 1) Or (ItemsShow >= MaxListLength) Then
  myItemsShow = IIf(MaxListLength > 8, MaxListLength - 8, ListCount)
 Else
  myItemsShow = ItemsShow
 End If
 Call PropertyChanged("NumberItemsToShow")
 Refresh
End Property

Public Property Get OfficeAppearance() As ComboOfficeAppearance
 '* English: Sets/Gets the office apperance.
 '* Español: Devuelve o establece la apariencia de Office.
 OfficeAppearance = myOfficeAppearance
End Property

Public Property Let OfficeAppearance(ByVal New_Apperance As ComboOfficeAppearance)
 myOfficeAppearance = New_Apperance
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("OfficeAppearance")
 Refresh
End Property

Public Property Get SelectBorderColor() As OLE_COLOR
 '* English: Sets/Gets the color of the border of the control when It has the focus.
 '* Español: Devuelve o establece el color del borde del control cuando el tenga el enfoque.
 SelectBorderColor = mySelectBorderColor
End Property

Public Property Let SelectBorderColor(ByVal New_Color As OLE_COLOR)
 mySelectBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectBorderColor")
 Refresh
End Property

Public Property Get SelectListBorderColor() As OLE_COLOR
 '* English: Sets/Gets the border color of the item selected in the list.
 '* Español: Devuelve o establece el color del borde del item seleccionado en la lista.
 SelectListBorderColor = mySelectListBorderColor
End Property

Public Property Let SelectListBorderColor(ByVal New_Color As OLE_COLOR)
 mySelectListBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectListBorderColor")
 Refresh
End Property

Public Property Get SelectListColor() As OLE_COLOR
 '* English: Sets/Gets the color of the item selected in the list.
 '* Español: Devuelve o establece el color del item seleccionado en la lista.
 SelectListColor = mySelectListColor
End Property

Public Property Let SelectListColor(ByVal New_Color As OLE_COLOR)
 mySelectListColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectListColor")
 Refresh
End Property

Public Property Get ShadowColorText() As OLE_COLOR
 '* English: Sets/Gets the text color of the shadow.
 '* Español: Devuelve o establece el color de la sombra del texto.
 ShadowColorText = myShadowColorText
End Property

Public Property Let ShadowColorText(ByVal New_Color As OLE_COLOR)
 myShadowColorText = ConvertSystemColor(New_Color)
 Call PropertyChanged("ShadowColorText")
 Refresh
End Property

Public Property Get Style() As ComboStyle
 '* English: Sets/Gets the style of the Combo.
 '* Español: Devuelve o establece el estilo del Combo.
 Style = myStyleCombo
End Property

Public Property Let Style(ByVal New_Style As ComboStyle)
 myStyleCombo = New_Style
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("Style")
 Refresh
End Property

Public Property Get Text() As String
 '* English: Sets/Gets the text of the selected item.
 '* Español: Devuelve o establece el texto del item seleccionado.
 Text = myText
End Property

Public Property Let Text(ByVal NewText As String)
 myText = NewText
 UserText = myText
 If (UserText = "") Then
  ItemFocus = -1
  picList.ListIndex = -1
 End If
 If (myStyleCombo = [Dropdown Combo]) Then txtCombo.Text = myText
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("Text")
End Property

Public Property Get XpAppearance() As ComboXpAppearance
 '* English: Sets the appearance in Xp Mode.
 '* Español: Establece la apariencia en modo Xp.
 XpAppearance = myXpAppearance
End Property

Public Property Let XpAppearance(ByVal New_Style As ComboXpAppearance)
 myXpAppearance = New_Style
 Call isEnabled(ControlEnabled)
 Call PropertyChanged("XpAppearance")
End Property

'********************************************************'
'* English: Subs and Functions of the Usercontrol.      *'
'* Español: Procedimientos y Funciones del Usercontrol. *'
'********************************************************'
Public Sub AddItem(ByVal Item As String, Optional ByVal ColorTextItem As OLE_COLOR = &HC56A31, Optional ByVal ImagePos As Integer = -1, Optional ByVal EnabledItem As Boolean = True, Optional ByVal ToolTipTextItem As String = "", Optional ByVal IndexItem As Long = -1, Optional ByVal ItemTag As String = "", Optional ByVal MouseIcon As StdPicture = Nothing, Optional ByVal SeparatorLine As Boolean = False, Optional ByVal TextShadow As Boolean = False)
 '* English: Add a new item to the list.
 '* Español: Agrega un nuevo item a la lista.
 If (Item = "") Then Item = " "
 sumItem = sumItem + 1
 ReDim Preserve ListContents(sumItem)
 If (IndexItem > 0) And (IndexItem < sumItem) And (NoFindIndex(IndexItem) = False) Then
  ListContents(sumItem).Index = IndexItem
 Else
  ListContents(sumItem).Index = sumItem
 End If
 ListContents(sumItem).Color = IIf(EnabledItem = True, ConvertSystemColor(ColorTextItem), ConvertSystemColor(DisabledColor))
 If (Len(Item) > Len(BigText)) Then BigText = Item
 ListContents(sumItem).Text = Item
 ListContents(sumItem).TextShadow = TextShadow
 ListContents(sumItem).Enabled = EnabledItem
 ListContents(sumItem).Index = IndexItem
 ListContents(sumItem).ToolTipText = ToolTipTextItem
 ListContents(sumItem).Tag = ItemTag
 ListContents(sumItem).ImagePos = ImagePos
 Set ListContents(sumItem).MouseIcon = MouseIcon
 ListContents(sumItem).SeparatorLine = SeparatorLine
 MaxListLength = sumItem
 Set picList.Font = UserControl.Font
 Call picList.AddItem(Item, sumItem, ImagePos, ListContents(sumItem).Color, EnabledItem, ToolTipTextItem, MouseIcon, SeparatorLine, TextShadow)
On Error GoTo myErr
 If Not (Images.ListImages(sumItem).Picture Is Nothing) Then
  isPicture = True
  Set ListContents(sumItem).Image = Images.ListImages(ImagePos).Picture
 End If
myErr:
On Error GoTo 0
 RaiseEvent TotalItems(sumItem)
End Sub

Private Sub APIFillRect(ByVal hDC As Long, ByRef RC As RECT, ByVal Color As Long)
 Dim NewBrush As Long
 
 '* English: The FillRect function fills a rectangle by using the specified brush. _
             This function includes the left and top borders, but excludes the right _
             and bottom borders of the rectangle.
 '* Español: Pinta el rectángulo de un objeto.
 NewBrush& = CreateSolidBrush(Color&)
 Call FillRect(hDC&, RC, NewBrush&)
 Call DeleteObject(NewBrush&)
End Sub

Private Sub APILine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)
 Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
 '* English: Use the API LineTo for Fast Drawing.
 '* Español: Pinta líneas de forma sencilla y rápida.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, x1, y1, PT)
 Call LineTo(UserControl.hDC, x2, y2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Function APIRectangle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional ByVal lColor As OLE_COLOR = -1) As Long
 Dim hPen As Long, hPenOld As Long
 Dim PT   As POINTAPI
 
 '* English: Paint a rectangle using API.
 '* Español: Pinta el rectángulo de un Objeto.
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

Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long) As Long
 '* English: Blend two colors in a 50%.
 '* Español: Mezclar dos colores al 50%.
 BlendColors = RGB(((lColor1 And &HFF) + (lColor2 And &HFF)) / 2, (((lColor1 \ &H100) And &HFF) + ((lColor2 \ &H100) And &HFF)) / 2, (((lColor1 \ &H10000) And &HFF) + ((lColor2 \ &H10000) And &HFF)) / 2)
End Function
        
Private Function CalcTextWidth(ByVal strCtlCaption As String, Optional ByVal strCtlCaption1 As String = "", Optional ByVal isW As Integer = 0) As String
 Dim lngMaxWidth  As Long, lngX As Long
 Dim lngTextWidth As Long
  
 '* English: Establishes the width of the text of the control.
 '* Español: Establece el ancho del texto del control.
 lngMaxWidth = UserControl.ScaleWidth - Int(UserControl.TextWidth(strCtlCaption) / 2) + 8
 lngTextWidth = UserControl.TextWidth(strCtlCaption) + isW
 If (strCtlCaption1 = "") Then strCtlCaption1 = strCtlCaption
 lngX = (Len(strCtlCaption) / 2)
 While (lngTextWidth > lngMaxWidth) And (lngX > 3)
  strCtlCaption1 = Mid$(strCtlCaption1, 1, lngX + 3) & IIf(Len(strCtlCaption1) = lngX, "", "...")
  lngTextWidth = UserControl.TextWidth(strCtlCaption1)
  lngX = lngX - 1
 Wend
 CalcTextWidth = strCtlCaption1
End Function
        
Public Sub ChangeItem(ByVal Index As Long, ByVal Item As String, Optional ByVal ColorTextItem As OLE_COLOR = &HC56A31, Optional ByVal ImagePos As Long = -1, Optional ByVal EnabledItem As Boolean = True, Optional ByVal ToolTipTextItem As String = "", Optional ByVal IndexItem As Long = -1, Optional ByVal ItemTag As String = "", Optional ByVal MouseIcon As StdPicture = Nothing, Optional ByVal SeparatorLine As Boolean = False, Optional ByVal TextShadow As Boolean = False)
 '* English: Modifies an item of the list.
 '* Español: Modifica un item de la lista.
 ListContents(Index).Color = IIf(EnabledItem = True, ColorTextItem, ShiftColorOXP(DisabledColor))
 ListContents(Index).Text = Item
 ListContents(Index).Enabled = EnabledItem
 If (IndexItem > 0) And (IndexItem < sumItem) And (NoFindIndex(IndexItem) = False) Then ListContents(Index).Index = IndexItem
 Set ListContents(Index).MouseIcon = MouseIcon
 ListContents(Index).SeparatorLine = SeparatorLine
 ListContents(Index).ToolTipText = ToolTipTextItem
 ListContents(Index).Tag = ItemTag
 ListContents(Index).ImagePos = ImagePos
On Error Resume Next
 Set ListContents(Index).Image = Images.ListImages(ImagePos).Picture
On Error GoTo 0
 ListContents(Index).TextShadow = TextShadow
 Call picList.ModifyItem(Index, Item, ImagePos)
End Sub
        
Public Sub Clear()
 '* English: Clear the list.
 '* Español: Borra toda la lista.
 sumItem = 0
 ReDim ListContents(0)
 Text = ""
 ItemFocus = -1
 IndexItemNow = -1
 ListIndex = -1
 isPicture = False
 RaiseEvent TotalItems(sumItem)
 picList.Clear
 Refresh
End Sub

Private Sub ControlsSubClasing()
 '* English: Add controls the Usercontrol.
 '* Español: Agrega controles al Usercontrol.
 Set picTemp = UserControl.Controls.Add("VB.PictureBox", "picTemp")
 picTemp.AutoRedraw = True
 picTemp.ScaleMode = vbPixels
 picTemp.AutoSize = True
 picTemp.TabStop = False
End Sub
        
Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 '* English: Convert Long to System Color.
 '* Español: Convierte un long en un color del sistema.
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function
        
Private Sub CreateImage(ByVal myPicture As StdPicture, ByVal ObjecthDC As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal Disabled As Boolean = False, Optional ByVal nHeight As Long = 16, Optional ByVal nWidth As Long = 16)
 Dim sTMPpathFName As String
 
 '* English: Draw the image in the Object.
 '* Español: Crea la imagen sobre el Objeto.
 Set picTemp.Picture = myPicture
 picTemp.BackColor = &HF0
 If (Disabled = False) Then
  Call PicDisabled(picTemp)
 Else
  sTMPpathFName = TempPathName + "\~ConvIconToBmp.tmp"
  Call SavePicture(picTemp.Image, sTMPpathFName)
  Set picTemp.Picture = LoadPicture(sTMPpathFName)
  Call Kill(sTMPpathFName)
 End If
 picTemp.Refresh
 Call CreateImageMask(picTemp, picTemp, &HF0)
 Call StretchBlt(ObjecthDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcAnd)
 Call CreateImageSprite(picTemp, picTemp, &HF0)
 Call StretchBlt(ObjecthDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcInvert)
End Sub

'* Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6077&lngWId=1. & _
   Thanks to David Peace Author of the code.
Private Function CreateImageMask(ByRef PicSrc As PictureBox, ByRef picDest As PictureBox, ByVal bColor As OLE_COLOR)
 Dim Looper  As Integer, Looper2 As Integer
 Dim bColor2 As OLE_COLOR

 picDest.Cls
 For Looper = 0 To PicSrc.Height
  picDest.Refresh
  For Looper2 = 0 To PicSrc.Width
   If (PicSrc.Point(Looper2, Looper) = bColor) Then
    bColor2 = RGB(255, 255, 255)
   Else
    bColor2 = RGB(0, 0, 0)
   End If
   Call SetPixel(picDest.hDC, Looper2, Looper, bColor2)
  Next
 Next
 picDest.Refresh
End Function

'* Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6077&lngWId=1. & _
   Thanks to David Peace Author of the code.
Private Function CreateImageSprite(ByRef PicSrc As PictureBox, ByRef picDest As PictureBox, ByVal bColor As OLE_COLOR)
 Dim Looper  As Integer, Looper2 As Integer
 Dim bColor2 As OLE_COLOR

 picDest.Cls
 For Looper = 0 To PicSrc.Height
  picDest.Refresh
  For Looper2 = 0 To PicSrc.Width
   If (PicSrc.Point(Looper2, Looper) = bColor) Then
    bColor2 = RGB(0, 0, 0)
   Else
    bColor2 = GetPixel(PicSrc.hDC, Looper2, Looper)
   End If
   SetPixel picDest.hDC, Looper2, Looper, bColor2
  Next
 Next
 picDest.Refresh
End Function

Private Function CreateMacOSXRegion() As Long
 Dim pPoligon(8) As POINTAPI, lW As Long, lh As Long
 
 '* English: Create a nonrectangular region for the MAC OS X Style.
 '* Español: Crea el Estilo MAC OS X.
 lW = UserControl.ScaleWidth
 lh = UserControl.ScaleHeight
 pPoligon(0).X = 0:      pPoligon(0).Y = 2
 pPoligon(1).X = 2:      pPoligon(1).Y = 0
 pPoligon(2).X = lW - 2: pPoligon(2).Y = 0
 pPoligon(3).X = lW:     pPoligon(3).Y = 2
 pPoligon(4).X = lW:     pPoligon(4).Y = lh - 5
 pPoligon(5).X = lW - 6: pPoligon(5).Y = lh
 pPoligon(6).X = 3:      pPoligon(6).Y = lh
 pPoligon(7).X = 0:      pPoligon(7).Y = lh - 3
 CreateMacOSXRegion = CreatePolygonRgn(pPoligon(0), 8, 1)
End Function

Private Sub DrawAppearance(Optional ByVal Style As ComboAppearance = 1, Optional ByVal m_State As Integer = 1)
 Dim isText    As String, isW As Integer, VImage As Integer
 Dim m_lRegion As Long, isH   As Integer, RText  As RECT
 
 '* English: Draw appearance of the control.
 '* Español: Dibuja la apariencia del control.
 Cls
 AutoRedraw = True
 FillStyle = 1
 m_StateG = m_State
 isH = 0
 If (Style <> 6) Then UserControl.BackColor = myBackColor
On Error Resume Next
 With txtCombo
  .Height = Abs(ScaleHeight / 2 - 7)
  .Top = Abs(ScaleHeight / 2 - 7)
  .ForeColor = IIf(Enabled = True, myNormalColorText, DisabledColor)
 On Error Resume Next
  Set .Font = g_Font
  If (myStyleCombo = 1) Then
   .Visible = False
  Else
   .Visible = True
  End If
 End With
 If (Height < 300) And (Style <> 11) Then
  Height = 300
 ElseIf (Height < 310) And (Style = 11) Then
  Height = 310
 ElseIf (Height > 600) Then
  If (Style = 12) Then
   Height = 300
  Else
   Height = 310
  End If
 End If
 If (Width < 840) Then Width = 840
 If (m_StateG <> 3) Then picList.Visible = False
 Select Case Style
  Case 1
   Call DrawOfficeButton(myOfficeAppearance)
  Case 2
   '* English: Style Windows 98.
   '* Español: Estilo Windows 98.
   Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, EDGE_SUNKEN)
   Call APIFillRect(UserControl.hDC, m_btnRect, GetSysColor(COLOR_BTNFACE))
   tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
   Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 3, EDGE_SUNKEN, EDGE_RAISED))
   Call DrawStandardArrow(m_btnRect, ArrowColor)
  Case 3
   '* English: Style Windows Xp.
   '* Español: Estilo Windows Xp.
   If (myXpAppearance = 1) Then     '* Aqua.
    tmpColor = &HB99D7F
   ElseIf (myXpAppearance = 2) Then '* Olive Green.
    tmpColor = &H94CCBC
   ElseIf (myXpAppearance = 3) Then '* Silver.
    tmpColor = &HA29594
   ElseIf (myXpAppearance = 4) Then '* TasBlue.
    tmpColor = &HF09F5F
   ElseIf (myXpAppearance = 5) Then '* Gold.
    tmpColor = &HBFE7F0
   ElseIf (myXpAppearance = 6) Then '* Blue.
    tmpColor = ShiftColorOXP(&HA0672F, 123)
   ElseIf (myXpAppearance = 7) Or (myXpAppearance = 0) Then '* Custom.
    If (m_StateG = 1) Then
     tmpColor = NormalBorderColor
    ElseIf (m_StateG = 2) Then
     tmpColor = HighLightBorderColor
    ElseIf (m_StateG = 3) Then
     tmpColor = SelectBorderColor
    End If
   End If
   Call DrawWinXPButton(myXpAppearance, tmpColor)
   If (myXpAppearance <> 0) Or (isFailedXP = True) Then
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, UserControl.BackColor)
   End If
  Case 4
   '* English: Style Soft.
   '* Español: Estilo Suavizado.
   Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER)
   Call APIFillRect(UserControl.hDC, m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(NormalBorderColor, 228), ShiftColorOXP(NormalBorderColor, 155)))
   tempBorderColor = GetSysColor(COLOR_BTNFACE)
   Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, GetSysColor(COLOR_BTNFACE))
   Call APILine(m_btnRect.Left - 1, m_btnRect.Top, m_btnRect.Left - 1, m_btnRect.Bottom, GetSysColor(COLOR_BTNFACE))
   Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 3, BDR_SUNKENOUTER, BDR_RAISEDINNER))
   Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(ArrowColor, 106), ArrowColor))
  Case 5
   Call DrawKDEButton
  Case 6
   '* English: Style MAC.
   '* Español: Estilo MAC.
   isH = 2
   Call DrawMacOSXCombo
  Case 7
   '* English: Style JAVA.
   '* Español: Estilo JAVA.
   tmpColor = ShiftColorOXP(NormalBorderColor, 52)
   tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
   Call DrawJavaBorder(0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
   Call APIFillRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 2, tmpColor, IIf(m_StateG <> -1, NormalBorderColor, ShiftColorOXP(NormalBorderColor, 192))))
   Call DrawJavaBorder(m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
   Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(ArrowColor, 166), ArrowColor))
  Case 8
   Call DrawExplorerBarButton(m_StateG)
  Case 9
   Dim tempPict As StdPicture, isWP As Integer
   
   '* English: Style User Picture.
   '* Español: Estilo Imagen de Usuario.
   Set tempPict = Nothing
   If (m_StateG = 1) Then
    Set tempPict = myNormalPictureUser
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    Set tempPict = myHighLightPictureUser
    tmpColor = HighLightBorderColor
   ElseIf (m_StateG = 3) Then
    Set tempPict = myFocusPictureUser
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
   Else
    Set tempPict = myDisabledPictureUser
    tmpColor = ShiftColorOXP(NormalBorderColor, 43)
   End If
   isH = 2
   If Not (tempPict Is Nothing) Then isWP = 19 Else isWP = 17
   Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - isWP, 0, UserControl.ScaleWidth - isWP, UserControl.ScaleHeight, GetLngColor(Parent.BackColor), False)
   Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth - isWP, UserControl.ScaleHeight, IIf(m_StateG <> -1, tmpColor, ShiftColorOXP(DisabledColor, 145)), True)
   If Not (tempPict Is Nothing) Then Call CreateImage(tempPict, UserControl.hDC, UserControl.ScaleWidth - 18, Abs(Int(UserControl.ScaleHeight / 2) - 11), True, 18, 17)
  Case 10
   '* English: Special Style.
   '* Español: Estilo especial con borde recortado.
   If (m_StateG = 1) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 75)
   ElseIf (m_StateG = 2) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 45)
   ElseIf (m_StateG = 3) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 15)
   Else
    tmpColor = ShiftColorOXP(&H0&, 237)
   End If
   Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 1, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 25), False)
   If (m_StateG = 1) Then
    tmpColor = ShiftColorOXP(&HC56A31, 143)
   ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
    tmpColor = ShiftColorOXP(&HC56A31, 113)
    tempBorderColor = tmpColor
   Else
    tmpColor = ShiftColorOXP(&H0&)
   End If
   Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5), True)
   tmpC2 = 12
   For tmpC1 = 2 To 5
    tmpC2 = tmpC2 + 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, BackColor)
   Next
   tmpC2 = 17
   tmpC3 = -2
   For tmpC1 = 5 To 2 Step -1
    tmpC2 = tmpC2 - 1
    tmpC3 = tmpC3 + 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
   Next
   Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
  Case 11
   '* English: Rounded Style.
   '* Español: Estilo Circular.
   isH = 2
   tempBorderColor = ShiftColorOXP(&H9F3000, 45)
   If (m_StateG = 1) Then
    iFor = &HCF989F
    tmpColor = &HA07F7F
    cValor = &HFFFFFF
   ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
    iFor = &H9F3000
    tmpColor = &HAF572F
    cValor = &HFFFFFF
   Else
    tmpColor = ShiftColorOXP(&H404040, 166)
    iFor = &HFFF8FF
    cValor = ShiftColorOXP(&H404040, 16)
   End If
   FillStyle = 0
   FillColor = iFor
   UserControl.Circle (m_btnRect.Left + 7, CInt(UserControl.ScaleHeight / 2)), 8, tmpColor
   UserControl.Circle (m_btnRect.Left + 7, CInt(UserControl.ScaleHeight / 2)), 7, &HFFFFFF
   Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   m_btnRect.Left = m_btnRect.Left - 5
   m_btnRect.Top = CInt(UserControl.ScaleHeight / 2) - 11
   UserControl.Line (m_btnRect.Left + 9, m_btnRect.Top + 8)-(m_btnRect.Left + 13, m_btnRect.Top + 12), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 10, m_btnRect.Top + 8)-(m_btnRect.Left + 13, m_btnRect.Top + 11), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 15, m_btnRect.Top + 8)-(m_btnRect.Left + 11, m_btnRect.Top + 12), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 14, m_btnRect.Top + 8)-(m_btnRect.Left + 11, m_btnRect.Top + 11), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 9, m_btnRect.Top + 12)-(m_btnRect.Left + 13, m_btnRect.Top + 16), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 10, m_btnRect.Top + 12)-(m_btnRect.Left + 13, m_btnRect.Top + 15), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 15, m_btnRect.Top + 12)-(m_btnRect.Left + 11, m_btnRect.Top + 16), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 14, m_btnRect.Top + 12)-(m_btnRect.Left + 11, m_btnRect.Top + 15), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
  Case 12
   Call DrawGradientButton(1)
  Case 13
   Call DrawGradientButton(2)
  Case 14
   Call DrawLightBlueButton
  Case 15
   '* English: Arrow Style.
   '* Español: Estilo Flecha.
   If (m_StateG = 1) Then
    cValor = GetLngColor(GradientColor1)
    iFor = GetLngColor(GradientColor2)
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
    iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
    tmpColor = HighLightBorderColor
   ElseIf (m_StateG = 3) Then
    cValor = GetLngColor(GradientColor1)
    iFor = GetLngColor(GradientColor2)
    tmpColor = SelectBorderColor
   Else
    cValor = GetLngColor(ShiftColorOXP(GradientColor1))
    iFor = GetLngColor(GradientColor2)
    tmpColor = ShiftColorOXP(&H0&)
   End If
   tempBorderColor = tmpColor
   Call DrawGradient(UserControl.hDC, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1, iFor, cValor, 1)
   Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
   cValor = UserControl.ScaleHeight / 2 + 1
   iFor = IIf(m_StateG = -1, &HC0C0C0, ArrowColor)
   For tmpColor = 7 To -2 Step -1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, cValor - (tmpColor / 2), UserControl.ScaleWidth - 7, cValor - (tmpColor / 2), IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
   Next
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 5, cValor, UserControl.ScaleWidth - 6, cValor, IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, cValor + 1, UserControl.ScaleWidth - 7, cValor + 1, IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 7, cValor + 2, UserControl.ScaleWidth - 8, cValor + 2, IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
  Case 16
   Call DrawNiaWBSSButton
  Case 17
   Call DrawRhombusButton
  Case 18
   Call DrawXpButton
  Case 19
   '* English: Ardent Style.
   '* Español: Estilo Ardent.
   If (m_StateG = 1) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 24)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 65)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = HighLightBorderColor
   ElseIf (m_StateG = 3) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 14)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
   Else
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2))
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = &HC0C0C0
   End If
   Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
   Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 19, UserControl.ScaleHeight - 2, ShiftColorOXP(cValor, 85))
   tmpC1 = 7
   tmpC2 = 4
   tmpC3 = ScaleHeight / 2 + 1
   For tmpColor = 6 To 2 Step -1
    tmpC1 = tmpC1 - 1
    tmpC2 = tmpC2 - 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC1, tmpC3 + tmpC2, UserControl.ScaleWidth - (tmpColor + 2), tmpC3 + tmpC2, IIf(m_StateG = -1, ShiftColorOXP(&HC0C0C0, 36), ArrowColor))
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC1, tmpC3 - 2 + tmpC2 - 1, UserControl.ScaleWidth - (tmpColor + 2), tmpC3 - 2 + tmpC2 - 1, IIf(m_StateG = -1, ShiftColorOXP(&HC0C0C0, 36), ArrowColor))
   Next
  Case 20
   Call DrawChocolateButton
  Case 21
   Call DrawButtonDownload
 End Select
 Call SetRect(m_btnRect, UserControl.ScaleWidth - 18, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)
 If (Style = 6) Then
  If (m_lRegion <> 0) Then Call DeleteObject(m_lRegion)
  m_lRegion = CreateMacOSXRegion
  Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
 Else
  m_lRegion = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
  Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
 End If
 If (ItemFocus > 0) And (ItemFocus <= ListCount) Then
  '* English: Sets the image of the current item.
  '* Español: Establece la imagen del item actual.
  picTemp.BackColor = ListColor
  If (ListContents(ItemFocus).Enabled = False) And (ThisFocus > 0) Then
   ItemFocus = ThisFocus
  End If
  If Not (ListContents(ItemFocus).Image Is Nothing) Then
   If (Style = 7) Or (Style = 6) Then VImage = 1 Else VImage = 0
   Call CreateImage(ListContents(ItemFocus).Image, UserControl.hDC, 5, Abs(Int(ScaleHeight / 2) - 7) - VImage, Enabled)
   cValor = 27
  Else
   cValor = 8
  End If
  isText = ListContents(ItemFocus).Text
  isW = 47 + isH
 Else
  isText = Text
  cValor = 8
  isW = 26 + isH
 End If
 txtCombo.Left = cValor
 If (myStyleCombo = 1) Then
  With UserControl
   .CurrentX = cValor
   .CurrentY = Int(UserControl.ScaleHeight / 2) - 7
   Set .Font = g_Font
   If (Enabled = False) Then
    Call SetTextColor(.hDC, DisabledColor)
   Else
    Call SetTextColor(.hDC, NormalColorText)
   End If
   isText = CalcTextWidth(myText, , 4)
   RText = m_btnRect
   RText.Left = cValor
   If (mWindowsNT = True) Then
    Call DrawTextW(.hDC, StrPtr(isText), Len(isText), RText, DT_VCENTER Or DT_LEFT Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
   Else
    Call DrawTextA(.hDC, isText, Len(isText), RText, DT_VCENTER Or DT_LEFT Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
   End If
  End With
 End If
 txtCombo.Width = Abs(ScaleWidth - isW) - 5
 txtCombo.BackColor = UserControl.BackColor
End Sub

Private Sub DrawButtonDownload()
 '* English: Draw Button Download appearance.
 '* Español: Crea la apariencia de un Botón de Descarga.
 cValor = IIf(m_StateG = -1, ShiftColorOXP(&H92603C), &H92603C)
 tempBorderColor = cValor
 tmpC3 = IIf(m_StateG = -1, ShiftColorOXP(&HE0C6AE), &HE0C6AE)
 If (m_StateG = 1) Or (m_StateG = 3) Then
  tmpC1 = &HBE8F63
  tmpC2 = &HE8DBCB
  tmpColor = ArrowColor
 ElseIf (m_StateG = 2) Then
  tmpC1 = ShiftColorOXP(&HBE8F63, 49)
  tmpC2 = ShiftColorOXP(&HE8DBCB, 49)
  tmpColor = ShiftColorOXP(ArrowColor, 89)
 Else
  tmpC1 = ShiftColorOXP(&HBE8F63)
  tmpC2 = ShiftColorOXP(&HE8DBCB)
  tmpColor = ShiftColorOXP(&HC0C0C0, 85)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 2, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, tmpC1, tmpC2, 1)
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, cValor)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(cValor, 5))
 Call DrawXpArrow(tmpColor)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, Int(UserControl.ScaleHeight / 2) + 4, UserControl.ScaleWidth - 6, Int(UserControl.ScaleHeight / 2) + 4, tmpColor)
 Call DrawShadow(tmpC3, tmpC3, False)
End Sub

Private Sub DrawChocolateButton()
 '* English: Chocolate Style.
 '* Español: Estilo Chocolate.
 cValor = IIf(m_StateG = -1, ShiftColorOXP(&H4A464B), &H4A464B)
 tempBorderColor = cValor
 tmpC3 = &HFFFFFF
 If (m_StateG = 1) Or (m_StateG = 3) Then
  tmpC1 = &H686567
  tmpC2 = ShiftColorOXP(&H292929, 89)
  tmpColor = &H0
 ElseIf (m_StateG = 2) Then
  tmpC1 = ShiftColorOXP(&H686567, 89)
  tmpC2 = ShiftColorOXP(&H292929, 178)
  tmpColor = ShiftColorOXP(&H0, 89)
 Else
  tmpC1 = ShiftColorOXP(&H838181)
  tmpC2 = ShiftColorOXP(&H292929)
  tmpColor = ShiftColorOXP(&H0)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 2, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, tmpC1, tmpC2, 2)
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, cValor)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(cValor, 5))
 Call DrawShadow(tmpC3, tmpColor, False)
 m_btnRect.Bottom = m_btnRect.Bottom / 2 + 4
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, m_btnRect.Bottom + 2, UserControl.ScaleWidth - 5, m_btnRect.Bottom + 2, tmpColor)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, m_btnRect.Bottom + 3, UserControl.ScaleWidth - 5, m_btnRect.Bottom + 3, tmpColor)
 For iFor = 4 To 7
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor - 1, m_btnRect.Bottom - iFor + 5, UserControl.ScaleWidth - iFor - 1, m_btnRect.Bottom - iFor + 5, tmpColor)
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor - 1, m_btnRect.Bottom + iFor, UserControl.ScaleWidth - (iFor + 1), m_btnRect.Bottom + iFor, tmpColor)
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor, m_btnRect.Bottom - iFor + 5, UserControl.ScaleWidth - (iFor + 2), m_btnRect.Bottom - iFor + 5, &HFFFFFF)
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor, m_btnRect.Bottom + iFor, UserControl.ScaleWidth - (iFor + 2), m_btnRect.Bottom + iFor, &HFFFFFF)
 Next
End Sub

Private Sub DrawCtlEdge(ByVal hDC As Long, ByVal X As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single, Optional ByVal Style As Long = EDGE_RAISED, Optional ByVal flags As Long = BF_RECT)
 Dim R As RECT
 
 '* English: The DrawEdge function draws one or more edges of rectangle. _
             using the specified coords.
 '* Español: Dibuja uno ó más bordes del rectángulo.
 With R
  .Left = X
  .Top = Y
  .Right = X + W
  .Bottom = Y + H
 End With
 Call DrawEdge(hDC, R, Style, flags)
End Sub

Private Sub DrawCtlEdgeByRect(ByVal hDC As Long, ByRef RT As RECT, Optional ByVal Style As Long = EDGE_RAISED, Optional ByVal flags As Long = BF_RECT)
 '* English: Draws the edge in a rect.
 '* Español: Colorea uno ó más bordes del rectángulo del Control.
 Call DrawEdge(hDC, RT, Style, flags)
End Sub

Private Sub DrawExplorerBarButton(ByVal m_StateG As Long)
 Dim isBackColor As OLE_COLOR
 
 '* English: Style ExplorerBar.
 '* Español: Estilo ExplorerBar.
 isBackColor = ShiftColorOXP(&HDEEAF0, 184)
 txtCombo.BackColor = isBackColor
 UserControl.BackColor = isBackColor
 Call DrawRectangleBorder(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, &HEAF3F7)
 If (m_StateG = 1) Then
  cValor = ShiftColorOXP(&HB6BFC3, 91)
  iFor = &HEAF3F7
  tmpColor = ShiftColorOXP(&HB6BFC3, 162)
 ElseIf (m_StateG = 2) Then
  cValor = ShiftColorOXP(&HB6BFC3, 31)
  iFor = &HDCEBF1
  tmpColor = ShiftColorOXP(&HB6BFC3, 132)
 ElseIf (m_StateG = 3) Then
  cValor = ShiftColorOXP(&HB6BFC3, 21)
  iFor = &HCEE3EC
  tmpColor = ShiftColorOXP(&HB6BFC3, 112)
  tempBorderColor = ShiftColorOXP(&HB6BFC3, 21)
 Else
  UserControl.BackColor = ShiftColorOXP(&HEAF3F7, 124)
  txtCombo.BackColor = UserControl.BackColor
  cValor = ShiftColorOXP(&HB6BFC3, 84)
  tmpC1 = ShiftColorOXP(&HEAF3F7, 124)
  iFor = ShiftColorOXP(&HEAF3F7, 123)
  tmpColor = ShiftColorOXP(&HB6BFC3, 132)
 End If
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, cValor)
 If (m_StateG = -1) Then Call DrawRectangleBorder(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpC1)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, iFor, False)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, UserControl.BackColor)
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 196), ArrowColor))
End Sub

Private Sub DrawGradient(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Direction As Integer)
 Dim Vert(1) As TRIVERTEX, gRect As GRADIENT_RECT

 '* English: Draw a gradient in the selected coords and hDC.
 '* Español: Dibuja el objeto en forma degradada.
 Call LongToRGB(Color1)
 With Vert(0)
  .X = X
  .Y = Y
  .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
  .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
  .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
  .Alpha = 1
 End With
 Call LongToRGB(Color2)
 With Vert(1)
  .X = x1
  .Y = y1
  .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
  .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
  .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
  .Alpha = 0
 End With
 gRect.UpperLeft = 0
 gRect.LowerRight = 1
 If (Direction = 1) Then
  Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)
 Else
  Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H)
 End If
End Sub

Private Sub DrawGradientButton(ByVal WhatGradient As Long)
 '* English: Draw a Vertical or Horizontal Gradient style appearance.
 '* Español: Dibuja la apariencia degradada bien sea vertical ó horizontal.
 If (m_StateG = 1) Then
  tmpColor = ShiftColorOXP(&HC56A31, 133)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HD8CEC5)
 ElseIf (m_StateG = 2) Then
  tmpColor = ShiftColorOXP(&HC56A31, 113)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HD6BEB5)
 ElseIf (m_StateG = 3) Then
  tmpColor = ShiftColorOXP(&HC56A31, 93)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HB3A29B)
  tempBorderColor = tmpColor
 Else
  tmpColor = CLng(ShiftColorOXP(&H0&))
  cValor = GetLngColor(&HC0C0C0)
  iFor = GetLngColor(&HFFFFFF)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, cValor, iFor, WhatGradient)
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor, True)
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawJavaBorder(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal lColorShadow As Long, ByVal lColorLight As Long, ByVal lColorBack As Long)
 '* English: Draw the edge with a JAVA style.
 '* Español: Dibuja el borde estilo JAVA.
 Call APIRectangle(UserControl.hDC, X, Y, W - 1, H - 1, lColorShadow)
 Call APIRectangle(UserControl.hDC, X + 1, Y + 1, W - 1, H - 1, lColorLight)
 Call SetPixel(UserControl.hDC, X, Y + H, lColorBack)
 Call SetPixel(UserControl.hDC, X + W, Y, lColorBack)
 Call SetPixel(UserControl.hDC, X + 1, Y + H - 1, BlendColors(lColorLight, lColorShadow))
 Call SetPixel(UserControl.hDC, X + W - 1, Y + 1, BlendColors(lColorLight, lColorShadow))
End Sub

Private Sub DrawKDEButton()
 '* English: Style KDE.
 '* Español: Estilo KDE.
 If (m_StateG = 1) Then
  tmpColor = NormalBorderColor
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 63))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 63))
 ElseIf (m_StateG = 2) Then
  tmpColor = HighLightBorderColor
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 127))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 127))
 ElseIf (m_StateG = 3) Then
  tmpColor = SelectBorderColor
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
 Else
  tmpColor = &HC0C0C0
  cValor = GetLngColor(&HFFFFFF)
  iFor = ShiftColorOXP(GetLngColor(&HC0C0C0), 45)
 End If
 tempBorderColor = tmpColor
 '* Español: Top Left.
 '* Español: Parte Superior Izquierda.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - 8, m_btnRect.Bottom - 8, iFor, cValor, 1)
 '* Español: Top Right.
 '* Español: Parte Inferior Derecha.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 8, m_btnRect.Top + 8, m_btnRect.Right, m_btnRect.Bottom, cValor, iFor, 1)
 '* Español: Bottom Right.
 '* Español: Parte Inferior Derecha.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 8, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom - 8, iFor, cValor, 1)
 '* Español: Bottom Left.
 '* Español: Parte Inferior Izquierda.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top + 8, m_btnRect.Right - 8, m_btnRect.Bottom, cValor, iFor, 1)
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor, True)
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawLightBlueButton()
 Dim PT      As POINTAPI, cx As Long, cy As Long
 Dim hPenOld As Long, hPen   As Long
   
 '* English: Style LightBlue.
 '* Español: Estilo LightBlue.
 If (m_StateG = 1) Or (m_StateG = 3) Then
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HA87057)
  tmpColor = &HA69182
  tempBorderColor = tmpColor
 ElseIf (m_StateG = 2) Then
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HCFA090)
  tmpColor = &HAF9080
 Else
  cValor = GetLngColor(&HFFFFFF)
  iFor = ShiftColorOXP(GetLngColor(&HA87057))
  tmpColor = ShiftColorOXP(&HA69182, 146)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1, cValor, iFor, 1)
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 If (m_StateG = 2) Then
  Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, &H53969F)
  Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 16, 1, 15, UserControl.ScaleHeight - 2, &H92C4D8)
  tmpColor = &H3EB4DE
 ElseIf (m_StateG <> -1) Then
  Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
  tmpColor = ArrowColor
 End If
 cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
 cy = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 + 2
 hPen = CreatePen(0, 1, IIf(m_StateG <> -1, tmpColor, ShiftColorOXP(&HC0C0C0, 97)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, cx - 3, cy - 1, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy - 1)
 Call LineTo(UserControl.hDC, cx, cy)
 Call LineTo(UserControl.hDC, cx - 2, cy)
 Call LineTo(UserControl.hDC, cx, cy + 2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
 hPen = CreatePen(0, 1, IIf(m_StateG <> -1, tmpColor, ShiftColorOXP(&HC0C0C0, 97)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 3
 Call MoveToEx(UserControl.hDC, cx - 4, cy - 3, PT)
 Call LineTo(UserControl.hDC, cx, cy - 3)
 Call LineTo(UserControl.hDC, cx - 2, cy - 5)
 Call LineTo(UserControl.hDC, cx - 3, cy - 4)
 Call LineTo(UserControl.hDC, cx - 1, cy - 3)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Sub DrawList(ByVal TopItem As Integer)
 Dim ListW As Long, Total As Long
 
 ListW = UserControl.TextWidth(BigText)
 With picList
  .ItemHeight = 20
  .ItemHeightAuto = False
  .VisibleRows = IIf(ListCount < NumberItemsToShow, ListCount, NumberItemsToShow)
  .Width = (50 + TextWidth(BigText)) * ScaleY(Screen.TwipsPerPixelY, vbTwips, vbPixels) + (picList.Width - picList.Width * ScaleX(Screen.TwipsPerPixelX, vbTwips, vbPixels))
  If (.Width < UserControl.ScaleWidth) Then .Width = UserControl.ScaleWidth
  .ItemOffset = 4
  If (.Font.Size > 8.25) Then .ItemHeightAuto = True: .ItemOffset = 0
  .BackNormal = myListColor
  .BackSelected = mySelectListColor
  If (TopItem < sumItem) Then .ListIndex = TopItem
  .FontSelected = myHighLightColorText
  .ListGradient = myListGradient
  .ScrollBarWidth = 18
  .BackSelectedG1 = MSSoftColor(myGradientColor1)
  .BackSelectedG2 = MSSoftColor(myGradientColor2)
  .HoverSelection = True
  .BackSelected = mySelectListColor
  .SelectListBorderColor = mySelectListBorderColor
  If (myAppearanceCombo = 8) Then
   .SelectBorderColor = tempBorderColor
  Else
   .SelectBorderColor = mySelectBorderColor
  End If
  .ShadowColorText = myShadowColorText
  If (isPicture = True) Then
   .ItemTextLeft = 22
  Else
   .ItemTextLeft = 4
  End If
  .WordWrap = False
  .Refresh
 End With
End Sub

Private Sub DrawMacOSXCombo()
 Dim PT      As POINTAPI, cy  As Long, cx     As Long, Color1 As Long, ColorG As Long
 Dim hPen    As Long, hPenOld As Long, Color2 As Long, Color3 As Long, ColorH As Long
 Dim Color4  As Long, Color5  As Long, Color6 As Long, Color7 As Long, ColorI As Long
 Dim Color8  As Long, Color9  As Long, ColorA As Long, ColorB As Long
 Dim ColorC  As Long, ColorD  As Long, ColorE As Long, ColorF As Long
 
 '* English: Draw the Mac OS X combo (this is a cool style!).
 '* Español: Dibujar el combo estilo Mac OS X (este es un estilo chevere).
 m_btnRect.Left = m_btnRect.Left - 4
 tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
 '* English: Button gradient top.
 ColorA = &HA0A0A0
 UserControl.BackColor = myBackColor
 If (m_StateG = 1) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 9)
  Color2 = ShiftColorOXP(&HDE8B45, 9)
  Color3 = ShiftColorOXP(&HDD873E, 9)
  Color4 = ShiftColorOXP(&HB33A01, 9)
  Color5 = ShiftColorOXP(&HE9BD96, 9)
  Color6 = ShiftColorOXP(&HB9B2AD, 9)
  Color7 = ShiftColorOXP(&H968A82, 9)
  Color8 = ShiftColorOXP(&HA25022, 9)
  Color9 = ShiftColorOXP(&HB8865E, 9)
  ColorB = ShiftColorOXP(&HDFBC86, 9)
  ColorC = ShiftColorOXP(&HFFBA77, 9)
  ColorD = ShiftColorOXP(&HE3D499, 9)
  ColorE = ShiftColorOXP(&HFFD996, 9)
  ColorF = ShiftColorOXP(&HE1A46D, 9)
  ColorG = ShiftColorOXP(&HCBA47B, 9)
  ColorH = ShiftColorOXP(&HDFDFDF, 9)
  ColorI = ShiftColorOXP(&HD0D0D0, 9)
 ElseIf (m_StateG = 2) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 89)
  Color2 = ShiftColorOXP(&HDE8B45, 89)
  Color3 = ShiftColorOXP(&HDD873E, 89)
  Color4 = ShiftColorOXP(&HB33A01, 99)
  Color5 = ShiftColorOXP(&HE9BD96, 109)
  Color6 = ShiftColorOXP(&HB9B2AD, 109)
  Color7 = ShiftColorOXP(&H968A82, 109)
  Color8 = ShiftColorOXP(&HA25022, 109)
  Color9 = ShiftColorOXP(&HB8865E, 109)
  ColorB = ShiftColorOXP(&HDFBC86, 109)
  ColorC = ShiftColorOXP(&HFFBA77, 109)
  ColorD = ShiftColorOXP(&HE3D499, 109)
  ColorE = ShiftColorOXP(&HFFD996, 109)
  ColorF = ShiftColorOXP(&HE1A46D, 109)
  ColorG = ShiftColorOXP(&HCBA47B, 109)
  ColorH = ShiftColorOXP(&HDFDFDF, 109)
  ColorI = ShiftColorOXP(&HD0D0D0, 109)
 ElseIf (m_StateG = 3) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 15)
  Color2 = ShiftColorOXP(&HDE8B45, 15)
  Color3 = ShiftColorOXP(&HDD873E, 15)
  Color4 = ShiftColorOXP(&HB33A01, 15)
  Color5 = ShiftColorOXP(&HE9BD96, 15)
  Color6 = ShiftColorOXP(&HB9B2AD, 15)
  Color7 = ShiftColorOXP(&H968A82, 15)
  Color8 = ShiftColorOXP(&HA25022, 15)
  Color9 = ShiftColorOXP(&HB8865E, 15)
  ColorB = ShiftColorOXP(&HDFBC86, 15)
  ColorC = ShiftColorOXP(&HFFBA77, 15)
  ColorD = ShiftColorOXP(&HE3D499, 15)
  ColorE = ShiftColorOXP(&HFFD996, 15)
  ColorF = ShiftColorOXP(&HE1A46D, 15)
  ColorG = ShiftColorOXP(&HCBA47B, 15)
  ColorH = ShiftColorOXP(&HDFDFDF, 15)
  ColorI = ShiftColorOXP(&HD0D0D0, 15)
 Else
  Color1 = ShiftColorOXP(&H808080, 195)
  Color2 = ShiftColorOXP(&H808080, 135)
  Color3 = ShiftColorOXP(&H808080, 135)
  Color4 = ShiftColorOXP(&H808080, 5)
  Color5 = Color1
  Color6 = GetLngColor(Parent.BackColor)
  Color7 = Color6
  Color8 = ShiftColorOXP(&H808080, 65)
  Color9 = Color6
  ColorA = Color6
  ColorB = Color4
  ColorC = Color4
  ColorD = Color4
  ColorE = Color4
  ColorF = Color4
  ColorG = Color4
  ColorH = Color6
  ColorI = Color6
 End If
 Call DrawVGradient(Color1, Color2, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 3)
 '* English: Button gradient bottom.
 Call DrawVGradient(Color3, Color1, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, UserControl.ScaleHeight / 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight * 2 / 3 - 4)
 '* English: Lines for the text area.
 Call APILine(2, 0, UserControl.ScaleWidth - 3, 0, &HA1A1A1)
 Call APILine(1, 0, 1, UserControl.ScaleHeight - 3, &HA1A1A1)
 '* English: Left shadow.
 If (m_StateG <> -1) Then
  Call DrawVGradient(ColorH, &HBBBBBB, 0, 0, 1, 3)
  Call DrawVGradient(&HBBBBBB, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
  Call DrawVGradient(ColorA, &HBBBBBB, 0, UserControl.ScaleHeight / 2, 1, UserControl.ScaleHeight / 2 - 5)
  Call DrawVGradient(&HBBBBBB, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
 Else
  Call DrawVGradient(ColorH, ColorH, 0, 0, 1, 3)
  Call DrawVGradient(ColorA, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
  Call DrawVGradient(ColorA, ColorA, 0, UserControl.ScaleHeight / 2, 1, UserControl.ScaleHeight / 2 - 5)
  Call DrawVGradient(ColorH, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
 End If
 '* English: Bottom shadows.
 Call APILine(1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, &H747474)
 Call APILine(1, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 2, &HA1A1A1)
 Call APILine(2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, &HDDDDDD)
 '* English: Lines for the button area.
 Call DrawVGradient(ColorB, Color3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, UserControl.ScaleHeight / 3)
 Call DrawVGradient(Color3, ColorB, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, UserControl.ScaleHeight / 3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, UserControl.ScaleHeight * 2 / 3 - 4)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 0, UserControl.ScaleWidth - 3, 0, Color4)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, 1, UserControl.ScaleWidth - 4, 1, Color5)
 '* English: Right shadow.
 Call DrawVGradient(ColorH, ColorI, UserControl.ScaleWidth - 1, 2, UserControl.ScaleWidth, 3)
 Call DrawVGradient(ColorI, ColorA, UserControl.ScaleWidth - 1, 3, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 6)
 Call DrawVGradient(ColorA, ColorI, UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 2 - 2, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 6)
 Call DrawVGradient(ColorI, ColorH, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 8, UserControl.ScaleWidth, 3)
 '* English: Layer1.
 Call DrawVGradient(Color4, Color3, UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 7)
 '* English: Layer2.
 Call DrawVGradient(Color4, ColorC, UserControl.ScaleWidth - 3, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6)
 '* English: Doted Area / 1-Bottom.
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4, ColorG)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 4, Color7)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5, ColorF)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 5, Color7)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 4, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 2, &HCACACA)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 2, &HBFBFBF)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 1, &HE4E4E4)
 '* English: Doted Area / 2-Botom
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 4, ColorD)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 5, ColorE)
 '* English: Doted Area / 3-Top.
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 0, IIf(m_StateG <> -1, &HA76E4A, ShiftColorOXP(&H808080, 55)))
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 0, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 1, Color8)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, 1, IIf(m_StateG <> -1, &HB3A49D, GetLngColor(Parent.BackColor)))
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, 1, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 1, Color8)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 2, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 3, Color8)
 '* English: Draw Twin Arrows.
 cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
 cy = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 - 1
 hPen = CreatePen(0, 1, IIf(m_StateG <> -1, &H0&, ShiftColorOXP(&H0&)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 '* English: Down Arrow.
 Call MoveToEx(UserControl.hDC, cx - 3, cy + 1, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy + 1)
 Call LineTo(UserControl.hDC, cx, cy + 2)
 Call LineTo(UserControl.hDC, cx - 2, cy + 2)
 Call LineTo(UserControl.hDC, cx - 2, cy + 3)
 Call LineTo(UserControl.hDC, cx, cy + 3)
 Call LineTo(UserControl.hDC, cx - 1, cy + 4)
 Call LineTo(UserControl.hDC, cx - 1, cy + 6)
 '* English: Up Arrow.
 Call MoveToEx(UserControl.hDC, cx - 3, cy - 2, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy - 2)
 Call LineTo(UserControl.hDC, cx, cy - 3)
 Call LineTo(UserControl.hDC, cx - 2, cy - 3)
 Call LineTo(UserControl.hDC, cx - 2, cy - 4)
 Call LineTo(UserControl.hDC, cx, cy - 4)
 Call LineTo(UserControl.hDC, cx - 1, cy - 5)
 Call LineTo(UserControl.hDC, cx - 1, cy - 7)
 '* English: Destroy PEN.
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
 '* English: Undo the offset.
 m_btnRect.Left = m_btnRect.Left + 4
End Sub

Private Sub DrawNiaWBSSButton()
 '* English: NiaWBSS Style.
 '* Español: Estilo NiaWBSS.
 If (m_StateG = 1) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = NormalBorderColor
 ElseIf (m_StateG = 2) Then
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
  tmpColor = HighLightBorderColor
 ElseIf (m_StateG = 3) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = SelectBorderColor
  tempBorderColor = tmpColor
 Else
  cValor = GetLngColor(ShiftColorOXP(GradientColor1))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2))
  tmpColor = ShiftColorOXP(DisabledColor, 156)
 End If
 Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
 tmpC1 = UserControl.ScaleHeight / 2 - 2
 tmpC2 = IIf(m_StateG = -1, DisabledColor, ArrowColor)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 2, tmpC1 - 1, UserControl.ScaleWidth - 12, tmpC1 - 1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 10, tmpC1 - 1, UserControl.ScaleWidth - 4, tmpC1 - 1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, tmpC1, UserControl.ScaleWidth - 10, tmpC1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 8, tmpC1, UserControl.ScaleWidth - 5, tmpC1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 4, tmpC1 + 1, UserControl.ScaleWidth - 6, tmpC1 + 1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 4, tmpC1 + 2, UserControl.ScaleWidth - 6, tmpC1 + 2, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 For tmpC3 = 3 To 6
  If (tmpC3 = 3) Or (tmpC3 = 4) Then
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 5, tmpC1 + tmpC3, UserControl.ScaleWidth - 7, tmpC1 + tmpC3, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
  Else
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, tmpC1 + tmpC3, UserControl.ScaleWidth - 8, tmpC1 + tmpC3, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
  End If
 Next
End Sub

Private Sub DrawOfficeButton(ByVal WhatOffice As ComboOfficeAppearance)
 Dim tmpRect As RECT
 
 '* English: Draw Office Style appearance.
 '* Español: Dibuja la apariencia de Office.
 tmpRect = m_btnRect
 Select Case WhatOffice
  Case 0
   '* English: Style Office Xp, appearance default.
   '* Español: Estilo Office Xp, apariencia por defecto.
   If (m_StateG = 1) Then
    '* English: Normal Color.
    '* Español: Color Normal.
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    '* English: Highlight Color.
    '* Español: Color de Selección MouseMove.
    tmpColor = HighLightBorderColor
    cValor = 185
   ElseIf (m_StateG = 3) Then
    '* English: Down Color.
    '* Español: Color de Selección MouseDown.
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
    cValor = 125
   Else
    '* English: Disabled Color.
    '* Español: Color deshabilitado.
    tmpColor = ConvertSystemColor(ShiftColorOXP(NormalBorderColor, 41))
   End If
   If (m_StateG > 1) Then
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tmpColor, B
    UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 14, UserControl.ScaleHeight - 2), ShiftColorOXP(tmpColor, cValor), BF
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), tmpColor, B
   Else
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tmpColor, B
    UserControl.Line (UserControl.ScaleWidth - 3, 2)-(UserControl.ScaleWidth - 13, UserControl.ScaleHeight - 3), tmpColor, BF
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), tmpColor, B
   End If
   Call DrawStandardArrow(m_btnRect, ArrowColor)
  Case 1
   '* English: Style Office 2000.
   '* Español: Estilo Office 2000.
   If (m_StateG = 1) Then
    '* English: Flat.
    '* Español: Normal.
    tmpColor = NormalBorderColor
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 13, 1, 12, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 175), False)
   ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
    '* English: Mouse Hover or Mouse Pushed.
    '* Español: Mouse presionado o MouseMove.
    If (m_StateG = 2) Then
     tmpColor = ShiftColorOXP(HighLightBorderColor)
    Else
     tmpColor = ShiftColorOXP(SelectBorderColor)
     tempBorderColor = tmpColor
    End If
    Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER)
    tmpRect.Left = tmpRect.Left + 4
    Call APIFillRect(UserControl.hDC, tmpRect, tmpColor)
    tmpRect.Left = tmpRect.Left - 1
    Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, tmpColor)
    Call APILine(tmpRect.Left, tmpRect.Top, tmpRect.Left, tmpRect.Bottom, tmpColor)
    If (m_StateG = 2) Then
     Call DrawCtlEdgeByRect(UserControl.hDC, tmpRect, BDR_RAISEDINNER)
    Else
     Call DrawCtlEdgeByRect(UserControl.hDC, tmpRect, BDR_SUNKENOUTER)
    End If
    m_btnRect.Left = m_btnRect.Left - 3
   Else
    '* English: Disabled control.
    '* Español: Control deshabilitado.
    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, ShiftColorOXP(&HC0C0C0, 36))
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 1, 17, UserControl.ScaleHeight - 2, myBackColor, False)
   End If
   tmpRect.Left = tmpRect.Left + 4
   Call DrawStandardArrow(tmpRect, IIf(m_StateG = -1, ShiftColorOXP(&HC0C0C0, 36), ArrowColor))
  Case 2
   '* English: Style Office 2003.
   '* Español: Estilo Office 2003.
   If (m_StateG <> -1) Then
    tmpC2 = GetSysColor(COLOR_WINDOW)
   Else
    tmpC2 = ShiftColorOXP(GetSysColor(COLOR_BTNFACE))
   End If
   tmpC1 = ArrowColor
   UserControl.BackColor = tmpC2
   txtCombo.BackColor = tmpC2
   tmpColor = GetSysColor(COLOR_HOTLIGHT)
   If (m_StateG = 1) Then
    cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), GetSysColor(29)), 109)
    iFor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_INACTIVECAPTIONTEXT), GetSysColor(COLOR_GRADIENTINACTIVECAPTION)))
   ElseIf (m_StateG = 2) Then
    cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), GetSysColor(29)), 170)
    iFor = cValor
   ElseIf (m_StateG = 3) Then
    cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), GetSysColor(29)), 140)
    iFor = cValor
   Else
    tmpC1 = GetSysColor(COLOR_GRAYTEXT)
    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpC1)
    txtCombo.ForeColor = tmpC1
    GoTo DrawNowArrow
   End If
   Call DrawGradient(UserControl.hDC, m_btnRect.Left + 4, tmpRect.Top - 1, tmpRect.Right + 1, tmpRect.Bottom + 1, iFor, cValor, 1)
   If (m_StateG = 2) Or (m_StateG = 3) Then
    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 15, 0, 17, UserControl.ScaleHeight, tmpColor, True)
    tempBorderColor = tmpColor
   End If
DrawNowArrow:
   Call DrawStandardArrow(tmpRect, tmpC1)
   myBackColor = tmpC2
 End Select
End Sub

Private Sub DrawRectangleBorder(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional ByVal SetBorder As Boolean = True)
 Dim hBrush As Long, TempRect As RECT

 '* English: Draw a rectangle.
 '* Español: Crea el rectángulo.
On Error Resume Next
 TempRect.Left = X
 TempRect.Top = Y
 TempRect.Right = X + Width
 TempRect.Bottom = Y + Height
 hBrush = CreateSolidBrush(Color)
 If (SetBorder = True) Then
  Call FrameRect(hDC, TempRect, hBrush)
 Else
  Call FillRect(hDC, TempRect, hBrush)
 End If
 Call DeleteObject(hBrush)
End Sub

Private Sub DrawRhombusButton()
 '* English: Rhombus Style.
 '* Español: Estilo Rombo.
 If (m_StateG = 1) Then
  tmpColor = ShiftColorOXP(NormalBorderColor, 25)
 ElseIf (m_StateG = 2) Then
  tmpColor = ShiftColorOXP(HighLightBorderColor, 25)
 ElseIf (m_StateG = 3) Then
  tmpColor = ShiftColorOXP(SelectBorderColor, 25)
 Else
  tmpColor = ShiftColorOXP(&H0&, 237)
 End If
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 1, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 25), False)
 If (m_StateG = 1) Then
  tmpColor = ShiftColorOXP(ArrowColor, 143)
 ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
  tmpColor = ShiftColorOXP(ArrowColor, 113)
  tempBorderColor = tmpColor
 Else
  tmpColor = ShiftColorOXP(&H0&)
 End If
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5), True)
 '* English: Left top border.
 '* Español: Borde Superior Izquierdo.
 tmpC2 = 12
 For tmpC1 = 2 To 5
  tmpC2 = tmpC2 + 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, BackColor)
 Next
 tmpC2 = 17
 tmpC3 = -2
 For tmpC1 = 5 To 2 Step -1
  tmpC2 = tmpC2 - 1
  tmpC3 = tmpC3 + 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
 Next
 '* English: Left bottom border.
 '* Español: Borde Inferior Izquierdo.
 tmpC2 = 17
 For tmpC1 = 3 To 1 Step -1
  tmpC2 = tmpC2 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, BackColor)
 Next
 tmpC2 = 12
 tmpC3 = 3
 For tmpC1 = 1 To 3
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, tmpColor)
 Next
 '* English: Right top border.
 '* Español: Borde Superior Derecho.
 tmpC2 = 0
 tmpC3 = 23
 For tmpC1 = 6 To 1 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, GetLngColor(Parent.BackColor))
 Next
 tmpC2 = 0
 tmpC3 = 17
 For tmpC1 = 6 To 1 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
 Next
 '* English: Right bottom border.
 '* Español: Borde Inferior Derecho.
 tmpC2 = 6
 For tmpC1 = 0 To 3
  tmpC2 = tmpC2 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 19, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, GetLngColor(Parent.BackColor))
 Next
 tmpC2 = 1
 tmpC3 = 16
 For tmpC1 = 3 To 0 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, UserControl.ScaleHeight - tmpC1 - 2, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 2, tmpColor)
 Next
 m_btnRect.Left = m_btnRect.Left + 1
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawShadow(ByVal iColor1 As Long, ByVal iColor2 As Long, Optional ByVal SoftColor As Boolean = True)
 '* English: Set a Shadow Border.
 '* Español: Coloca un borde con sombra.
 tmpC2 = 15
 If (SoftColor = True) Then
  tmpC3 = 178
  iFor = 10
 Else
  tmpC3 = 0
  iFor = 0
 End If
 For tmpC1 = 1 To 16
  tmpC2 = tmpC2 - 1
  '* Horizontal Top Border.
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, 1, UserControl.ScaleWidth - tmpC1, 1, ShiftColorOXP(iColor1, tmpC3))
  '* Horizontal Bottom Border.
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - tmpC1, UserControl.ScaleHeight - 2, IIf(m_StateG = -1, ShiftColorOXP(iColor2, tmpC3), ShiftColorOXP(iColor2, iFor)))
  If (SoftColor = True) Then
   tmpC3 = tmpC3 - 5
   iFor = iFor + 5
  End If
 Next
 m_btnRect.Bottom = m_btnRect.Bottom - 11
 If (SoftColor = True) Then
  tmpC3 = 128
  iFor = 70
 End If
 For tmpC1 = 0 To 12
  '* Vertical Left Border.
  Call APILine(m_btnRect.Left + 1, m_btnRect.Top + tmpC1 - 1, m_btnRect.Left + 1, m_btnRect.Bottom + tmpC1 - 1, ShiftColorOXP(iColor1, tmpC3))
  '* Vertical Right Border.
  Call APILine(UserControl.ScaleWidth - 2, m_btnRect.Top + tmpC1 - 1, UserControl.ScaleWidth - 2, m_btnRect.Bottom + tmpC1 - 1, IIf(m_StateG = -1, ShiftColorOXP(iColor2, tmpC3), ShiftColorOXP(iColor2, iFor)))
  If (SoftColor = True) Then
   tmpC3 = tmpC3 + 5
   iFor = iFor - 5
  End If
 Next
End Sub

Private Sub DrawStandardArrow(ByRef RT As RECT, ByVal lColor As Long)
 Dim PT   As POINTAPI, hPenOld As Long, cx As Long
 Dim hPen As Long, cy          As Long
 
 '* English: Draw the standard arrow in a Rect.
 '* Español: Dibuje la flecha normal en un Rect.
 If (AppearanceCombo = 1) And (OfficeAppearance = 1) Or (AppearanceCombo = 10) Or (AppearanceCombo = 17) Then
  hPen = 1
 ElseIf ((OfficeAppearance = 2) Or (OfficeAppearance = 0)) And (AppearanceCombo = 1) Then
  hPen = 2
 End If
 cx = RT.Left + (RT.Right - RT.Left) - (7 - hPen)
 cy = RT.Top + (RT.Bottom - RT.Top) / 2
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, cx - 3, cy - 1, PT)
 Call LineTo(UserControl.hDC, cx + 1, cy - 1)
 Call LineTo(UserControl.hDC, cx, cy)
 Call LineTo(UserControl.hDC, cx - 2, cy)
 Call LineTo(UserControl.hDC, cx, cy + 2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Sub DrawTextEx(ByVal hDC As Long, ByVal isText As String, ByVal nRctIndex As Long, ByVal isFormat As Variant)
 '* English: Set the text of the object.
 '* Español: Crea el texto sobre el objeto.
 
 '*************************************************************************
 '* Draws the text with Unicode support based on OS version.              *
 '* Thanks to Richard Mewett.                                             *
 '*************************************************************************
 If (mWindowsNT = True) Then
  Call DrawTextW(hDC, StrPtr(isText), Len(isText), m_TextRct(nRctIndex), isFormat)
 Else
  Call DrawTextA(hDC, isText, Len(isText), m_TextRct(nRctIndex), isFormat)
 End If
End Sub

Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT) As Boolean
 Dim hTheme  As Long '* hTheme Handle.
 Dim lResult As Long '* Temp Variable.
 
 '* If a error occurs then or we are not running XP or the visual style is Windows Classic.
On Error GoTo NoXP
 '* Get out hTheme Handle.
 hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
 '* Did we get a theme handle?.
 If (hTheme) Then
  '* Yes! Draw the control Background.
  lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
  '* If drawing was successful, return true, or false If not.
  DrawTheme = IIf(lResult, False, True)
 Else
  '* No, we couldn't get a hTheme, drawing failed.
  DrawTheme = False
 End If
 '* Close theme.
 Call CloseThemeData(hTheme)
 '* Exit the function now.
 Exit Function
NoXP:
 '* An Error was detected, drawing Failed.
 DrawTheme = False
End Function

Private Sub DrawVGradient(ByVal lEndColor As Long, ByVal lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal x2 As Long, ByVal y2 As Long)
 Dim dR As Single, dG As Single, dB As Single, ni As Long
 Dim sR As Single, sG As Single, Sb As Single
 Dim eR As Single, eG As Single, eB As Single
 
 '* English: Draw a Vertical Gradient in the current hDC.
 '* Español: Dibuja un degradado en forma vertical.
 sR = (lStartColor And &HFF)
 sG = (lStartColor \ &H100) And &HFF
 Sb = (lStartColor And &HFF0000) / &H10000
 eR = (lEndColor And &HFF)
 eG = (lEndColor \ &H100) And &HFF
 eB = (lEndColor And &HFF0000) / &H10000
 dR = (sR - eR) / y2
 dG = (sG - eG) / y2
 dB = (Sb - eB) / y2
 For ni = 0 To y2
  Call APILine(X, Y + ni, x2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
 Next
End Sub

Private Sub DrawWinXPButton(ByVal XpAppearance As ComboXpAppearance, ByVal tmpColor As OLE_COLOR)
 Dim tmpXPAppearance   As ComboXpAppearance, isState As Integer
 Dim bDrawThemeSuccess As Boolean, tmpRect           As RECT
 
 '* English: This Sub Draws the XpAppearance Button.
 '* Español: Este procedimiento dibuja el Botón estilo XP.
 isFailedXP = False
 If (XpAppearance = 0) Then
  '* Draw the XP Themed Style.
  isState = IIf(m_StateG < 0, 4, m_StateG)
  Call SetRect(tmpRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
  bDrawThemeSuccess = DrawTheme("Edit", 2, isState, tmpRect)
  Call SetRect(tmpRect, m_btnRect.Left - 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
  bDrawThemeSuccess = DrawTheme("ComboBox", 1, isState, tmpRect)
  If (bDrawThemeSuccess = True) Then
   Exit Sub
  Else '* If themed failed, then use the Next Style.
   tmpXPAppearance = 7 '* If failed, use custom colors.
   isFailedXP = True
   GoTo noUxThemed
  End If
 Else
  tmpXPAppearance = XpAppearance
 End If
noUxThemed:
 If (tmpXPAppearance = 7) And (m_StateG <> -1) Then
  UserControl.BackColor = BackColor
 ElseIf (m_StateG <> -1) Then
  UserControl.BackColor = GetSysColor(COLOR_WINDOW)
 Else
  UserControl.BackColor = &HE5ECEC
 End If
 Call APIRectangle(UserControl.hDC, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, IIf(m_StateG <> -1, tmpColor, &HC2C9C9))
 Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, GetSysColor(COLOR_WINDOW))
 Select Case tmpXPAppearance
  Case 1
   '* English: Style WinXp Aqua.
   '* Español: Estilo WinXp Aqua.
   cValor = &H85614D
   tempBorderColor = &HC56A31
   tmpC2 = &HB99D7F
   If (m_StateG = 1) Then
    tmpC3 = &HF5C8B3
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = ShiftColorOXP(&HF5C8B3, 58)
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HF9A477
    tmpColor = &HFFFFFF
   End If
  Case 2
   '* English: Style WinXp Olive Green.
   '* Español: Estilo WinXp Olive Green.
   cValor = &HFFFFFF
   tempBorderColor = &H668C7D
   tmpC2 = &H94CCBC
   If (m_StateG = 1) Then
    tmpC3 = &H8BB4A4
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = &HA7D7CA
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = &H80AA98
    tmpColor = &HFFFFFF
   End If
  Case 3
   '* English: Style WinXp Silver.
   '* Español: Estilo WinXp Silver.
   tempBorderColor = &HA29594
   cValor = &H48483E
   tmpC2 = &HA29594
   If (m_StateG = 1) Then
    tmpC3 = &HDACCCB
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = ShiftColorOXP(&HDACCCB, 58)
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HE5D1CF
    tmpColor = &HFFFFFF
   End If
  Case 4
   '* English: Style WinXp TasBlue.
   '* Español: Estilo WinXp TasBlue.
   tempBorderColor = &HF09F5F
   cValor = ShiftColorOXP(&H703F00, 58)
   tmpC2 = &HF09F5F
   If (m_StateG = 1) Then
    tmpC3 = &HF0AF70
    tmpColor = &HFFE7CF
   ElseIf (m_StateG = 2) Then
    tmpC3 = ShiftColorOXP(&HF0BF80, 58)
    tmpColor = &HFFEFD0
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HF09F5F
    tmpColor = &HFFEFD0
   End If
  Case 5
   '* English: Style WinXp Gold.
   '* Español: Estilo WinXp Gold.
   tempBorderColor = &HBFE7F0
   cValor = ShiftColorOXP(&H6F5820, 45)
   tmpC2 = &HBFE7F0
   If (m_StateG = 1) Then
    tmpC3 = ShiftColorOXP(&HCFFFFF, 54)
    tmpColor = &HBFF0FF
   ElseIf (m_StateG = 2) Then
    tmpC3 = &HBFEFFF
    tmpColor = ShiftColorOXP(&HCFFFFF, 58)
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HCFFFFF
    tmpColor = &HBFE8FF
   End If
  Case 6
   '* English: Style WinXp Blue.
   '* Español: Estilo WinXp Blue.
   tempBorderColor = ShiftColorOXP(&HA0672F, 123)
   cValor = &H6F5820
   tmpC2 = ShiftColorOXP(&HA0672F, 123)
   If (m_StateG = 1) Then
    tmpC3 = &HEFF0F0
    tmpColor = &HF0F7F0
   ElseIf (m_StateG = 2) Then
    tmpC3 = &HF0F8FF
    tmpColor = &HF0F7F0
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HF1946E
    tmpColor = &HEEC2B4
   End If
  Case 7
   '* English: Style WinXp Custom.
   '* Español: Estilo WinXp Custom.
   tempBorderColor = SelectBorderColor
   cValor = ArrowColor
   If (m_StateG = 1) Then
    tmpC3 = NormalBorderColor
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = HighLightBorderColor
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = SelectBorderColor
    tmpColor = &HFFFFFF
   End If
   tmpC2 = tmpC3
 End Select
 If (m_StateG = -1) Then
  tmpColor = &HE5ECEC
  tmpC3 = m_btnRect.Bottom - m_btnRect.Top
  tmpC1 = m_btnRect.Bottom - 1
  For iFor = 3 To tmpC1
   Call APILine(m_btnRect.Left + 1, tmpC3 - iFor + 3, m_btnRect.Right - 1, tmpC3 - iFor + 3, tmpColor)
  Next
  tmpC1 = ShiftColorOXP(&HC2C9C9, 19)
 Else
  tmpC1 = tmpC2
  Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, tmpColor, tmpC3, 1)
 End If
 Call APIRectangle(hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, tmpC1)
 Call DrawXpArrow(IIf(m_StateG = -1, &HC2C9C9, cValor))
End Sub

Private Sub DrawXpArrow(Optional ByVal iColor3 As OLE_COLOR = &H0)
 '* English: Draw The XP Style Arrow.
 '* Español: Dibuja la flecha estilo Xp.
 tmpC1 = m_btnRect.Right - m_btnRect.Left
 tmpC2 = m_btnRect.Bottom - m_btnRect.Top + 1
 tmpC1 = m_btnRect.Left + tmpC1 / 2 + 1
 tmpC2 = m_btnRect.Top + tmpC2 / 2
 If (iColor3 = &H0) Then iColor3 = ArrowColor
 Call APILine(tmpC1 - 5, tmpC2 - 2, tmpC1, tmpC2 + 3, iColor3)
 Call APILine(tmpC1 - 4, tmpC2 - 2, tmpC1, tmpC2 + 2, iColor3)
 Call APILine(tmpC1 - 4, tmpC2 - 3, tmpC1, tmpC2 + 1, iColor3)
 Call APILine(tmpC1 + 3, tmpC2 - 2, tmpC1 - 2, tmpC2 + 3, iColor3)
 Call APILine(tmpC1 + 2, tmpC2 - 2, tmpC1 - 2, tmpC2 + 2, iColor3)
 Call APILine(tmpC1 + 2, tmpC2 - 3, tmpC1 - 2, tmpC2 + 1, iColor3)
End Sub

Private Sub DrawXpButton()
 '* English: Additional Xp Style.
 '* Español: Estilo Xp Adicional.
 If (m_StateG = 1) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = NormalBorderColor
 ElseIf (m_StateG = 2) Then
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
  tmpColor = HighLightBorderColor
 ElseIf (m_StateG = 3) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = SelectBorderColor
  tempBorderColor = tmpColor
 Else
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 120))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 120))
  tmpColor = DisabledColor
 End If
 Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
 Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
 Call APILine(1, 1, UserControl.ScaleWidth - 18, 1, IIf(m_StateG = -1, ShiftColorOXP(DisabledColor, 98), ShiftColorOXP(tmpColor, 168)))
 Call DrawXpArrow(IIf(m_StateG <> -1, ArrowColor, tmpColor))
 Call DrawShadow(GradientColor1, &H646464)
 Call APILine(1, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 18, UserControl.ScaleHeight - 2, IIf(m_StateG = -1, ShiftColorOXP(DisabledColor, 98), ShiftColorOXP(tmpColor, 168)))
 Call APILine(1, 1, 1, UserControl.ScaleHeight - 2, IIf(m_StateG = -1, ShiftColorOXP(DisabledColor, 98), ShiftColorOXP(tmpColor, 168)))
End Sub

Private Sub Espera(ByVal Segundos As Single)
 Dim ComienzoSeg As Single, FinSeg As Single
 
 '* English: Wait a certain time.
 '* Español: Esperar un determinado tiempo.
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 Do While FinSeg > Timer
  DoEvents
  If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
 Loop
End Sub

Public Function FindItemText(ByVal Text As String, Optional ByVal Compare As StringCompare = 0) As Long
 Dim i As Long, RText As Long, SText As Long
 
 '* English: Search Text in the list and return the position.
 '* Español: Busca una cadena dentro de la lista y devuelve su posición en la misma.
 FindItemText = -1
 If (Text = "") Or (Compare < 0) Or (Compare > 2) Then Exit Function
 For i = 1 To sumItem
  If (Compare = 0) Then
   If (InStr(1, UCase$(ListContents(i).Text), UCase$(Text), vbTextCompare) <> 0) Then
    FindItemText = i
    Exit For
   End If
  ElseIf (Compare = 1) Then
   If (UCase$(Text) = UCase$(ListContents(i).Text)) Then
    FindItemText = i
    Exit For
   End If
  Else
   RText = AscW(Text)
   SText = AscW(ListContents(i).Text)
   If (Text = ListContents(i).Text) Then
    FindItemText = i
    Exit For
   'ElseIf (RText = SText) Then
   ' FindItemText = i
   ' Exit For
   End If
  End If
 Next
End Function

Public Function FindItemTag(ByVal Text As String, Optional ByVal Compare As StringCompare = 0) As Long
 Dim i As Long, RText As Long, SText As Long
 
 '* English: Search Text in the list and return the position.
 '* Español: Busca una cadena dentro de la lista y devuelve su posición en la misma.
 FindItemTag = -1
 If (Text = "") Or (Compare < 0) Or (Compare > 2) Then Exit Function
 For i = 1 To sumItem
  If (Compare = 0) Then
   If (InStr(1, UCase$(ListContents(i).Tag), UCase$(Text), vbTextCompare) <> 0) Then
    FindItemTag = i
    Exit For
   End If
  ElseIf (Compare = 1) Then
   If (UCase$(Text) = UCase$(ListContents(i).Tag)) Then
    FindItemTag = i
    Exit For
   End If
  Else
   RText = AscW(Text)
   SText = AscW(ListContents(i).Tag)
   If (Text = ListContents(i).Tag) Then
    FindItemTag = i
    Exit For
   'ElseIf (RText = SText) Then
   ' FindItemTag = i
   ' Exit For
   End If
  End If
 Next
End Function

Public Function GetControlVersion() As String
 '* English: Control Version.
 '* Español: Version del Control.
 GetControlVersion = Version & " © " & Year(Now)
End Function

Private Function GetLngColor(ByVal Color As Long) As Long
 '* English: The GetSysColor function retrieves the current color of the specified display element. Display elements are the parts of a window and the Windows display that appear on the system display screen.
 '* Español: Recupera el color actual del elemento de despliegue especificado.
 If (Color And &H80000000) Then
  GetLngColor = GetSysColor(Color And &H7FFFFFFF)
 Else
  GetLngColor = Color
 End If
End Function

Private Function InFocusControl(ByVal ObjecthWnd As Long) As Boolean
 Dim mPos As POINTAPI, oRect As RECT
 
 '* English: Verifies if the mouse is on the object or if one makes clic outside of him.
 '* Español: Verifica si el mouse se encuentra sobre el objeto ó si se hace clic fuera de él.
 Call GetCursorPos(mPos)
 Call GetWindowRect(ObjecthWnd, oRect)
 UserControl.MousePointer = myMousePointer
 '* English: Set MouseIcon only drop down list.
 '* Español: Coloca el icono del mouse únicamente donde se expande ó retrae la lista.
 'If (mPos.X > oRect.Left + (UserControl.ScaleWidth - 18)) And (mPos.X < oRect.Right) Then
 ' Set UserControl.MouseIcon = myMouseIcon
 'Else
 ' Set UserControl.MouseIcon = Nothing
 'End If
 If (mPos.X >= oRect.Left) And (mPos.X <= oRect.Right) And (mPos.Y >= oRect.Top) And (mPos.Y <= oRect.Bottom) Then
  InFocusControl = True
 End If
End Function

Private Sub isEnabled(ByVal isTrue As Boolean)
 '* English: Shows the state of Enabled or Disabled of the Control.
 '* Español: Muestra el estado de Habilitado ó Deshabilitado del Control.
 If (isTrue = True) Then
  Call DrawAppearance(myAppearanceCombo, 1)
 Else
  Call DrawAppearance(myAppearanceCombo, -1)
 End If
End Sub

Public Sub ItemEnabled(ByVal ListIndex As Long, ByVal ValueItem As Boolean)
 '* English: Sets the Enabled/disabled property in an Item.
 '* Español: Habilita o Deshabilita un Item.
On Error GoTo myErr:
 ListContents(ListIndex).Enabled = ValueItem
 Exit Sub
myErr:
End Sub

Public Function List(ByVal ListIndex As Long) As String
 '* English: Show one item of the list.
 '* Español: Muestra un elemento de la lista.
On Error Resume Next
 ItemFocus = ListIndex
 picList.ListIndex = ItemFocus - 1
 List = ListContents(ListIndex).Text
 Call isEnabled(ControlEnabled)
On Error GoTo 0
End Function

Public Function LoadRecordSet(ByRef mRecordset As Object, ByVal IDField As Integer, ByVal ItemIDTag As Integer) As Boolean
 Dim tmpVal  As String, i As Long
 Dim tmpVal1 As String
 
 '* English:
 '* Español:
 LoadRecordSet = False
On Error GoTo myErr
 While Not (mRecordset.EOF = True)
  tmpVal = IIf(IsNull(mRecordset(IDField).Value) = True, "", mRecordset(IDField).Value)
  tmpVal1 = IIf(IsNull(mRecordset(ItemIDTag).Value) = True, "", mRecordset(ItemIDTag).Value)
  Call AddItem(tmpVal, myNormalColorText, , , , , tmpVal1)
  mRecordset.MoveNext
 Wend
 LoadRecordSet = True
 Exit Function
myErr:
 LoadRecordSet = False
End Function

Private Sub LongToRGB(ByVal lColor As Long)
 '* English: Convert a Long to RGB format.
 '* Español: Convierte un Long en formato RGB.
 RGBColor.Red = lColor And &HFF
 RGBColor.Green = (lColor \ &H100) And &HFF
 RGBColor.Blue = (lColor \ &H10000) And &HFF
End Sub

Private Function MSSoftColor(ByVal lColor As Long) As Long
 Dim lRed  As Long, lGreen As Long, lb As Long
 Dim lBlue As Long, lr     As Long, lg As Long
 
 '* English: Set a soft color.
 '* Español: Devuelve un color suave.
 lr = (lColor And &HFF)
 lg = ((lColor And 65280) \ 256)
 lb = ((lColor) And 16711680) \ 65536
 lRed = (76 - Int(((lColor And &HFF) + 32) \ 64) * 19)
 lGreen = (76 - Int((((lColor And 65280) \ 256) + 32) \ 64) * 19)
 lBlue = (76 - Int((((lColor And &HFF0000) \ &H10000) + 32) / 64) * 19)
 MSSoftColor = RGB(lr + lRed, lg + lGreen, lb + lBlue)
End Function

Private Function NoFindIndex(ByVal Index As Long) As Boolean
 Dim i As Long
 
 '* English: Search if the Index has not been assigned.
 '* Español: Busca si ya no se ha asignado este Index.
 NoFindIndex = False
 For i = 1 To sumItem
  If (ListContents(i).Index = Index) Then NoFindIndex = True: Exit For
 Next
End Function

Public Sub OrderList(Optional ByVal Order As Integer = 1)
 Dim N As Long, i As Long, j As Long
 
 '* English: Order the list with the search method (I Exchange).
 '* Español: Ordena la lista con el método de búsqueda (Intercambio).
 If (Order <> 1) And (Order <> 2) Then Exit Sub
 ReDim OrderListContents(0)
 N = UBound(ListContents)
 For i = 1 To N
  ReDim Preserve OrderListContents(i)
  OrderListContents(i).Color = ListContents(i).Color
  OrderListContents(i).Enabled = ListContents(i).Enabled
  Set OrderListContents(i).Image = ListContents(i).Image
  OrderListContents(i).Index = ListContents(i).Index
  Set OrderListContents(i).MouseIcon = ListContents(i).MouseIcon
  OrderListContents(i).SeparatorLine = ListContents(i).SeparatorLine
  OrderListContents(i).Tag = ListContents(i).Tag
  OrderListContents(i).Text = ListContents(i).Text
  OrderListContents(i).ToolTipText = ListContents(i).ToolTipText
 Next
 i = 1
 For i = 1 To N
  For j = (i + 1) To N
   Select Case Order
    Case 1: If (OrderListContents(j).Text < OrderListContents(i).Text) Then Call SetInfo(i, j)
    Case 2: If (OrderListContents(j).Text > OrderListContents(i).Text) Then Call SetInfo(i, j)
   End Select
  Next
 Next
 ReDim ListContents(0)
 Call picList.Clear
 For i = 1 To N
  ReDim Preserve ListContents(i)
  ListContents(i).Color = OrderListContents(i).Color
  ListContents(i).Enabled = OrderListContents(i).Enabled
  Set ListContents(i).Image = OrderListContents(i).Image
  ListContents(i).Index = OrderListContents(i).Index
  Set ListContents(i).MouseIcon = OrderListContents(i).MouseIcon
  ListContents(i).SeparatorLine = OrderListContents(i).SeparatorLine
  ListContents(i).Tag = OrderListContents(i).Tag
  ListContents(i).Text = OrderListContents(i).Text
  ListContents(i).ToolTipText = OrderListContents(i).ToolTipText
  Call picList.AddItem(ListContents(i).Text)
 Next
 ReDim OrderListContents(0)
End Sub

Private Sub PicDisabled(ByRef picTo As PictureBox, Optional ByVal isGrayIcon As Boolean = True)
 Dim sTMPpathFName As String, lFlags As Long
 
 '* English: Disables a image.
 '* Español: Deshabilita la imagen.
 Select Case picTo.Picture.Type
  Case vbPicTypeBitmap
   lFlags = DST_BITMAP
  Case vbPicTypeIcon
   lFlags = DST_ICON
  Case Else
   lFlags = DST_COMPLEX
 End Select
 If Not (picTo.Picture Is Nothing) And (isGrayIcon = False) Then
  Call DrawState(picTo.hDC, 0, 0, picTo.Picture, 0, 0, 0, picTo.ScaleWidth, picTo.ScaleHeight, lFlags Or DSS_DISABLED)
 ElseIf Not (picTo.Picture Is Nothing) Then
  Call RenderIconGrayscale(picTo.hDC, picTo.Picture.Handle, 0, 0, picTo.ScaleWidth, picTo.ScaleHeight)
 End If
 sTMPpathFName = TempPathName + "\~ConvIconToBmp.tmp"
 Call SavePicture(picTo.Image, sTMPpathFName)
 Set picTo.Picture = LoadPicture(sTMPpathFName)
 Call Kill(sTMPpathFName)
 picTo.Refresh
End Sub

Public Sub RemoveItem(ByVal Index As Long)
 Dim TempList() As PropertyCombo, sCount As Long
 Dim Count      As Long, TempCount       As Long
 
 '* English: Delete a Item from the list.
 '* Español: Elimina un elemento de la lista.
On Error GoTo myErr
 If (ListCount = 0) Or (Index > ListCount) Then Exit Sub
 If (sumItem > 0) Then sumItem = Abs(sumItem - 1)
 For Count = 1 To picList.ListCount
  If (Index <> Count) Then
   sCount = sCount + 1
   ReDim Preserve TempList(sCount)
   TempList(sCount).Color = ListContents(Count).Color
   TempList(sCount).Enabled = ListContents(Count).Enabled
   Set TempList(sCount).Image = ListContents(Count).Image
   TempList(sCount).ImagePos = ListContents(Count).ImagePos
   TempList(sCount).Index = sCount
   TempList(sCount).Tag = ListContents(Count).Tag
   TempList(sCount).Text = ListContents(Count).Text
   TempList(sCount).ToolTipText = ListContents(Count).ToolTipText
   Set TempList(sCount).MouseIcon = ListContents(Count).MouseIcon
   TempList(sCount).SeparatorLine = ListContents(Count).SeparatorLine
   TempList(sCount).TextShadow = ListContents(Count).TextShadow
  End If
 Next
 TempCount = sCount
 sCount = 0
 ReDim ListContents(0)
 Call picList.Clear
 For Count = 1 To TempCount
  sCount = sCount + 1
  ReDim Preserve ListContents(sCount)
  ListContents(sCount).Color = TempList(Count).Color
  ListContents(sCount).Enabled = TempList(Count).Enabled
  Set ListContents(sCount).Image = TempList(Count).Image
  ListContents(sCount).ImagePos = TempList(Count).ImagePos
  ListContents(sCount).Index = TempList(Count).Index
  ListContents(sCount).Tag = TempList(Count).Tag
  ListContents(sCount).Text = TempList(Count).Text
  ListContents(sCount).ToolTipText = TempList(Count).ToolTipText
  Set ListContents(sCount).MouseIcon = TempList(Count).MouseIcon
  ListContents(sCount).SeparatorLine = TempList(Count).SeparatorLine
  ListContents(sCount).TextShadow = TempList(Count).TextShadow
  Call picList.AddItem(ListContents(sCount).Text, ListContents(sCount).Index, ListContents(sCount).ImagePos, ListContents(sCount).Color, ListContents(sCount).Enabled, ListContents(sCount).ToolTipText, ListContents(sCount).MouseIcon, ListContents(sCount).SeparatorLine, ListContents(sCount).TextShadow)
 Next
 MaxListLength = Abs(MaxListLength - 1)
On Error Resume Next
 If (myText = ListContents(MaxListLength + 1).Text) Then
  ListIndex = -1
  ItemFocus = -1
  Text = ""
  Call isEnabled(ControlEnabled)
 ElseIf (ListIndex = Index) Then
  ListIndex = -1
  ItemFocus = -1
  Text = ""
 End If
 RaiseEvent TotalItems(sumItem)
On Error GoTo 0
 Exit Sub
myErr:
End Sub

Public Sub SetImageList(ByRef ImgList As Object)
 Set Images = ImgList
 Call picList.SetImageList(Images)
End Sub

Private Sub SetInfo(ByVal i As Long, ByVal j As Long)
 Dim Temp As Variant
 
 '* English: Reorders the values.
 '* Español: Reordena los valores.
 Temp = OrderListContents(i).Color
 OrderListContents(i).Color = OrderListContents(j).Color
 OrderListContents(j).Color = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Enabled
 OrderListContents(i).Enabled = OrderListContents(j).Enabled
 OrderListContents(j).Enabled = Temp
 '*******************************************************************'
 Set Temp = OrderListContents(i).Image
 Set OrderListContents(i).Image = OrderListContents(j).Image
 Set OrderListContents(j).Image = Temp
 '*******************************************************************'
 Set Temp = OrderListContents(i).MouseIcon
 Set OrderListContents(i).MouseIcon = OrderListContents(j).MouseIcon
 Set OrderListContents(j).MouseIcon = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).SeparatorLine
 OrderListContents(i).SeparatorLine = OrderListContents(j).SeparatorLine
 OrderListContents(j).SeparatorLine = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Tag
 OrderListContents(i).Tag = OrderListContents(j).Tag
 OrderListContents(j).Tag = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Text
 OrderListContents(i).Text = OrderListContents(j).Text
 OrderListContents(j).Text = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).ToolTipText
 OrderListContents(i).ToolTipText = OrderListContents(j).ToolTipText
 OrderListContents(j).ToolTipText = Temp
End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
 Dim cRed   As Long, cBlue  As Long
 Dim Delta  As Long, cGreen As Long

 '* English: Shift a color.
 '* Español: Devuelve un Color con menos intensidad.
 cBlue = ((theColor \ &H10000) Mod &H100)
 cGreen = ((theColor \ &H100) Mod &H100)
 cRed = (theColor And &HFF)
 Delta = &HFF - Base
 cBlue = Base + cBlue * Delta \ &HFF
 cGreen = Base + cGreen * Delta \ &HFF
 cRed = Base + cRed * Delta \ &HFF
 If (cRed > 255) Then cRed = 255
 If (cGreen > 255) Then cGreen = 255
 If (cBlue > 255) Then cBlue = 255
 ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue
End Function

   '********************************'
   '*  Extracted of KPD-Team 1998  *'
   '*  URL: http://www.allapi.net  *'
   '*  E-Mail: KPDTeam@Allapi.net  *'
   '********************************'
Private Function TempPathName() As String
 Dim strTemp As String
 
 '* English: Returns the name of the temporary directory of Windows.
 '* Español: Devuelve el nombre del directorio temporal de Windows.
 strTemp = String$(100, Chr$(0)) '* Create a buffer.
 Call GetTempPath(100, strTemp)  '* Get the temporary path.
 '* Strip the rest of the buffer.
 TempPathName = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
End Function

'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    FreeLibrary hMod
  End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      TrackMouseEvent tme
    Else
      TrackMouseEventComCtl tme
    End If
  End If
End Sub

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(i)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = 0 Then                                                  'If the element is free...
        zData(i) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = uMsg Then                                               'If the message is found...
        zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
 If (ControlEnabled = False) Then Exit Sub
 Select Case lng_hWnd
  Case UserControl.hWnd, txtCombo.hWnd
   Select Case uMsg
    Case WM_THEMECHANGED
     Call DrawAppearance
    Case WM_MOUSEMOVE
     If Not (bInCtrl = True) Then
      bInCtrl = True
      Call TrackMouseLeave(lng_hWnd)
      Call DrawAppearance(myAppearanceCombo, 2)
      Call InFocusControl(lng_hWnd)
      RaiseEvent MouseEnter
     End If
    Case WM_MOUSELEAVE
     If (NoShow = False) And ((InFocusControl(txtCombo.hWnd) = False) Or (InFocusControl(picList.hWnd) = False)) Then
      bInCtrl = False
      Call DrawAppearance(myAppearanceCombo, 1)
     End If
     RaiseEvent MouseLeave
    Case WM_MOUSEWHEEL
     '* Based on original code of fred.cpp. Please see _
        the web site http://mx.geocities.com/fred_cpp/ _
        isexplorerbar.htm.
     If (isScroll = True) Then
      If (wParam = &H780000) Then
       picList.TopIndex = picList.TopIndex - 1
      ElseIf (wParam = &HFF880000) Then
       picList.TopIndex = picList.TopIndex + 1
      End If
     End If
   End Select
   Call InFocusControl(lng_hWnd)
  Case Extender.Parent.hWnd
   Select Case uMsg
    Case WM_MOUSELEAVE, WM_MOVING, WM_SIZING, _
         WM_EXITSIZEMOVE, WM_RBUTTONDOWN, _
         WM_MBUTTONDOWN, WM_LBUTTONDOWN, WM_ACTIVATE, _
         WM_NCLBUTTONDOWN
     NoShow = False
     bInCtrl = False
     picList.Visible = False
     Call DrawAppearance(myAppearanceCombo, 1)
    'Case WM_MOUSEMOVE, WM_MOUSELEAVE
     'NoShow = bInCtrl
   End Select
 End Select
End Sub
