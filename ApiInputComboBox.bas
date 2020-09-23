Attribute VB_Name = "APIInputBoxComboBox"
'Module name: APIInputBox & ComboBox
'Author: M.C
'Created: 09 may.02

'main Api Window Creating  learned from:
'API Form by Joseph Huntley found on PSC               *

'ver 1.0
'first bugy attempt
'ver 2.0 (excelency acchived)
' in ver 1.0 nothing could be typed in text portion of combo box
' this is now OK
' in ver 1.0 wierd thing happend to Num lock, caps lock on keyboard
' this is now ok
' added ability to swallow selection on ENTER key pressed
' horizontal and vertical scroll added as/and only if needed
' autotype ability added
' this last one vas especialy hard to figure out as API combobox doesn't
' get WM_CHAR message
'? button added, unfortunately doesn't work.




'tons of declares, constants and stuff
'a lot of them not needed in this project

'**************************************************************
'tons of green stuff - as evidence of trial - error method, lol
'**************************************************************

Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowText Lib "user632" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hwnd As Long
End Type

Const WM_MOVE = &H3
Const WM_SETCURSOR = &H20
Const WM_NCPAINT = &H85
Const WM_COMMAND = &H111

Const SWP_FRAMECHANGED = &H20
Const GWL_EXSTYLE = -20

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2

Public Const CW_USEDEFAULT = &H80000000

Public Const ES_MULTILINE = &H4&

' Window styles
Public Const WS_ACTIVECAPTION = &H1
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000         ' WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_CONTEXTHELP = &H400&
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_LAYOUTRTL = &H400000
Public Const WS_EX_LEFT = &H0&
Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_LTRREADING = &H0&
Public Const WS_EX_MDICHILD = &H40&
Public Const WS_EX_NOACTIVATE = &H8000000
Public Const WS_EX_NOINHERITLAYOUT = &H100000
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Public Const WS_EX_RIGHT = &H1000&
Public Const WS_EX_RIGHTSCROLLBAR = &H0&
Public Const WS_EX_RTLREADING = &H2000&
Public Const WS_EX_STATICEDGE = &H20000

Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_TABSTOP = &H10000
Public Const WS_GROUP = &H20000
Public Const WS_GT = (WS_GROUP Or WS_TABSTOP)
Public Const WS_HSCROLL = &H100000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000

'buttonstyles
Private Const BS_3STATE = &H5&
Private Const BS_AUTO3STATE = &H6&
Private Const BS_AUTOCHECKBOX = &H3&
Private Const BS_AUTORADIOBUTTON = &H9&
Private Const BS_BITMAP = &H80&
Private Const BS_BOTTOM = &H800&
Private Const BS_CENTER = &H300&
Private Const BS_CHECKBOX = &H2&
Private Const BS_DEFPUSHBUTTON = &H1&
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERN8X8 = 8
Private Const BS_DIBPATTERNPT = 6
Private Const BS_FLAT = &H8000&
Private Const BS_GROUPBOX = &H7&
Private Const BS_HATCHED = 2
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_ICON = &H40&
Private Const BS_INDEXED = 4
Private Const BS_LEFT = &H100&
Private Const BS_LEFTTEXT = &H20&
Private Const BS_MONOPATTERN = 9
Private Const BS_MULTILINE = &H2000&
Private Const BS_NOTIFY = &H4000&
Private Const BS_OWNERDRAW = &HB&
Private Const BS_PATTERN = 3
Private Const BS_PATTERN8X8 = 7
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_PUSHLIKE = &H1000&
Private Const BS_RADIOBUTTON = &H4&
Private Const BS_RIGHT = &H200&
Private Const BS_RIGHTBUTTON = BS_LEFTTEXT
Private Const BS_SOLID = 0
Private Const BS_TEXT = &H0&
Private Const BS_TOP = &H400&
Private Const BS_USERBUTTON = &H8&
Private Const BS_VCENTER = &HC00&


Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_BTNFACE = 15
Public Const COLOR_3DFACE = COLOR_BTNFACE
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_3DSHADOW = COLOR_BTNSHADOW
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADD = 712
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BLUE = 708
Public Const COLOR_BLUEACCEL = 728
Public Const COLOR_BOX1 = 720


Public Const COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_CURRENT = 709
Public Const COLOR_CUSTOM1 = 721
Public Const COLOR_DESKTOP = COLOR_BACKGROUND
Public Const COLOR_ELEMENT = 716
Public Const COLOR_GRADIENTACTIVECAPTION = 27
Public Const COLOR_GRADIENTINACTIVECAPTION = 28
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_GREEN = 707
Public Const COLOR_GREENACCEL = 727
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_HOTLIGHT = 26
Public Const COLOR_HUE = 703
Public Const COLOR_HUEACCEL = 723
Public Const COLOR_HUESCROLL = 700
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_INFOBK = 24
Public Const COLOR_INFOTEXT = 23
Public Const COLOR_LUM = 705
Public Const COLOR_LUMACCEL = 725
Public Const COLOR_LUMSCROLL = 702
Public Const COLOR_MATCH_VERSION = &H200
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_MIX = 719
Public Const COLOR_NO_TRANSPARENT = &HFFFFFFFF
Public Const COLOR_PALETTE = 718
Public Const COLOR_RAINBOW = 710
Public Const COLOR_RED = 706
Public Const COLOR_REDACCEL = 726
Public Const COLOR_SAMPLES = 717
Public Const COLOR_SAT = 704
Public Const COLOR_SATACCEL = 724
Public Const COLOR_SATSCROLL = 701
Public Const COLOR_SAVE = 711
Public Const COLOR_SCHEMES = 715
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_SOLID = 713
Public Const COLOR_SOLID_LEFT = 730
Public Const COLOR_SOLID_RIGHT = 731
Public Const COLOR_TUNE = 714
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8
Public Const COLORMATCHTOTARGET_EMBEDED = &H1
Public Const COLORMGMTCAPS = 121
Public Const COLORMGMTDLGORD = 1551
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const COLOROKSTRINGA = "commdlg_ColorOK"
Public Const COLOROKSTRINGW = "commdlg_ColorOK"
Public Const COLORONCOLOR = 3
Public Const COLORRES = 108


Public Const WM_DESTROY = &H2
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_GETTEXT = &HD
Public Const WM_SETTEXT = &HC

Public Const IDC_ARROW = 32512&

Public Const IDI_APPLICATION = 32512&
Public Const IDI_ASTERISK = 32516&
Public Const IDI_CLASSICON_OVERLAYFIRST = 500
Public Const IDI_CLASSICON_OVERLAYLAST = 502
Public Const IDI_CONFLICT = 161
Public Const IDI_DISABLED_OVL = 501
Public Const IDI_HAND = 32513&
Public Const IDI_ERROR = IDI_HAND
Public Const IDI_EXCLAMATION = 32515&
Public Const IDI_FORCED_OVL = 502

Public Const IDI_INFORMATION = IDI_ASTERISK
Public Const IDI_PROBLEM_OVL = 500
Public Const IDI_QUESTION = 32514&
Public Const IDI_RESOURCE = 159
Public Const IDI_RESOURCEFIRST = 159
Public Const IDI_RESOURCELAST = 161
Public Const IDI_RESOURCEOVERLAYFIRST = 161
Public Const IDI_RESOURCEOVERLAYLAST = 161
Public Const IDI_WARNING = IDI_EXCLAMATION
Public Const IDI_WINLOGO = 32517
Public Const IDIGNORE = 5

Private Const HWND_TOPMOST = -1


Public Const GWL_WNDPROC = (-4)

Public Const SW_SHOWNORMAL = 1
Private Const SW_HIDE = 0

Private Const LB_SETHORIZONTALEXTENT = &H194

Public Const MB_OK = &H0&
Public Const MB_ICONEXCLAMATION = &H30&


Public Const gClassName = "MyClassName"
Public Const gAppName = "My Window Caption"


'ComboBox styles
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_DISABLENOSCROLL = &H800&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_LOWERCASE = &H4000&
Public Const CBS_NOINTEGRALHEIGHT = &H400&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_SIMPLE = &H1&
Public Const CBS_SORT = &H100&
Public Const CBS_UPPERCASE = &H2000&

Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_DIR = &H145
Public Const CB_ERR = (-1)
Public Const CB_ERRSPACE = (-2)
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETDROPPEDCONTROLRECT = &H152
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_GETEDITSEL = &H140
Public Const CB_GETEXTENDEDUI = &H156
Public Const CB_GETHORIZONTALEXTENT = &H15D
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETLOCALE = &H15A
Public Const CB_GETTOPINDEX = &H15B
Public Const CB_INITSTORAGE = &H161
Public Const CB_INSERTSTRING = &H14A
Public Const CB_LIMITTEXT = &H141
Public Const CB_MSGMAX = &H15B
Public Const CB_MULTIPLEADDSTRING = &H163
Public Const CB_OKAY = 0
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_SETEDITSEL = &H142
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_SETHORIZONTALEXTENT = &H15E
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_SETLOCALE = &H159
Public Const CB_SETTOPINDEX = &H15C
Public Const CB_SHOWDROPDOWN = &H14F

Public Const WM_UPDATEUISTATE = &H128

Public Const EM_SETSEL = &HB1

Private Const BM_CLICK = &HF5

Public OKButtonOldProc As Long ''Will hold address of the old window proc for the button
Public CancelButtonOldProc As Long
Public ComboBoxOldProc As Long

Public FormWindowHwnd As Long, OKButtonHwnd As Long, gEditHwnd As Long, TextHwnd As Long, CancelButtonHwnd As Long, ComboBoxHwnd As Long, HelpButtonHwnd As Long  ''You don't necessarily need globals, but if you're planning to gettext and stuff, then you're gona have to store the hwnds.

'and some public dims
Dim InputBoxCaption As String
Dim InputBoxText As String
Dim AddItems() As Variant  'array
Dim i As Integer
Public IBCBSelectedItem As Variant 'this will hold whatever you select in combo box
Public avoid As Boolean
Public CharCountNew As Integer
Public CharCountOld As Integer
Public FirstTime307Passed As Boolean ' used in combobox procedure
Public textportion
Public trouble As Boolean
Private WHook As Long


Public Sub CreateComboBoxInputBox(Caption As String, IBText As String, InputAddItems() As Variant)
InputBoxCaption = Caption
InputBoxText = IBText
AddItems = InputAddItems
CharCountOld = 0
Main
End Sub
Public Sub Main()

   Dim wMsg As Msg

   ''Call procedure to register window classname. If false, then exit.
   If RegisterWindowClass = False Then Exit Sub
    
      ''Create window
      If CreateWindows Then
         ''Loop will exit when WM_QUIT is sent to the window.
         Do While GetMessage(wMsg, 0&, 0&, 0&)
            ''TranslateMessage takes keyboard messages and converts
            ''them to WM_CHAR for easier processing.
            Call TranslateMessage(wMsg)
            ''Dispatchmessage calls the default window procedure
            ''to process the window message. (WndProc)
            Call DispatchMessage(wMsg)
         Loop
      End If

    Call UnregisterClass(gClassName$, App.hInstance)


End Sub

Public Function RegisterWindowClass() As Boolean

    Dim wc As WNDCLASS
    
    ''Registers our new window with windows so we
    ''can use our classname.
    
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnwndproc = GetAddress(AddressOf WndProc) ''Address in memory of default window procedure.
    wc.hInstance = App.hInstance
    wc.hIcon = LoadIcon(0&, IDI_APLICATION) ''Default application icon
    wc.hCursor = LoadCursor(0&, IDC_ARROW) ''Default arrow
    wc.hbrBackground = COLOR_WINDOW ''Default a color for window.
    wc.lpszClassName = gClassName$

    RegisterWindowClass = RegisterClass(wc) <> 0
    
End Function
Public Function CreateWindows() As Boolean
  
    ''Create main window.
    FormWindowHwnd = CreateWindowEx(0, gClassName$, InputBoxCaption, WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME, (Screen.Width / Screen.TwipsPerPixelX / 2) - 150, (Screen.Height / Screen.TwipsPerPixelY / 2) - 80, 300, 160, 0, 0, App.hInstance, ByVal 0&)
    'FormWindowHwnd& = CreateWindowEx(0&, gClassName$, gAppName$, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 208, 150, 0&, 0&, App.hInstance, ByVal 0&)
    'Create OK button
    OKButtonHwnd = CreateWindowEx(0, "Button", "OK", WS_CHILD, 230, 10, 60, 25, FormWindowHwnd, 0, App.hInstance, ByVal 0&)
    'Create Cancel button
    CancelButtonHwnd = CreateWindowEx(0&, "button", "Cancel", WS_CHILD, 230, 45, 60, 25, FormWindowHwnd, 0&, App.hInstance, 0&)
    'OKButtonHwnd = CreateWindowEx(0&, "Button", "Click Here", WS_CHILD, 58, 90, 85, 25, FormWindowHwnd&, 0&, App.hInstance, 0&)
    ''Create textbox with a border (WS_EX_CLIENTEDGE) and make it multi-line (ES_MULTILINE)
    'gEditHwnd& = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "This is the edit control." & vbCrLf & "As you can see, it's multiline.", WS_CHILD Or ES_MULTILINE, 0&, 0&, 200, 80, FormWindowHwnd&, 0&, App.hInstance, 0&)
    
    'Create 'label'
    TextHwnd = CreateWindowEx(&H0, "static", InputBoxText, WS_CHILD, 5, 10, 200, 60, FormWindowHwnd, 0&, App.hInstance, 0&)
    
    'Create Combo
    ComboBoxHwnd = CreateWindowEx(0&, "combobox", "H", CBS_DROPDOWN Or CBS_HASSTRINGS Or CBS_SORT Or CBS_AUTOHSCROLL Or WS_CHILD Or WS_VSCROLL Or WS_HSCROLL, 5, 80, 285, 100, FormWindowHwnd, 0, App.hInstance, 0&)
    
    'Create Help button
    HelpButtonHwnd = CreateWindowEx(0&, "Button", "?", WS_CHILD, 0, -10, 14, 14, FormWindowHwnd, 0&, App.hInstance, ByVal 0&)
    
    

    'Tmp = SendMessage(ComboBoxHwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

    'Windows are all hidden, so show them.
    Call ShowWindow(FormWindowHwnd, SW_SHOWNORMAL)
    Call ShowWindow(OKButtonHwnd, SW_SHOWNORMAL)
    Call ShowWindow(CancelButtonHwnd, SW_SHOWNORMAL)
    'Call ShowWindow(gEditHwnd&, SW_SHOWNORMAL)
    Call ShowWindow(TextHwnd, SW_SHOWNORMAL)
    Call ShowWindow(ComboBoxHwnd, SW_SHOWNORMAL)
    Call ShowWindow(HelpButtonHwnd, SW_SHOWNORMAL)
    
    'help button in caption bar
    Dim FormWindowRect As Rect
    GetWindowRect FormWindowHwnd, FormWindowRect
    'procedure for help button
    'WHook = SetWindowsHookEx(4, AddressOf HelpButtonProc, 0, App.ThreadID)
    SetParent HelpButtonHwnd, GetParent(FormWindowHwnd)
    SetWindowPos HelpButtonHwnd, 0, FormWindowRect.Right - 40, FormWindowRect.Top + 6, 17, 14, SWP_FRAMECHANGED
    
    'Dim FormWindowRect As Rect
    'GetWindowRect FormWindowHwnd, FormWindowRect
    'SetWindowPos HelpButtonHwnd, 0, FormWindowRect.Right - 75, FormWindowRect.Top + 6, 17, 14, SWP_FRAMECHANGED
     
    
    'Call SetParent(HelpButtonHwnd, GetParent(FormWindowHwnd))
   'Initialize the window hooking for the button
    'WHook = SetWindowsHookEx(4, AddressOf HelpButtonProc, 0, App.ThreadID)
    'Call SetWindowLong(HelpButtonHwnd, GWL_EXSTYLE, &H80)
    'Call SetParent(HelpButtonHwnd, GetParent(FormWindowHwnd))
    
    'At this point the following pairs of lines must be written for each api created Window
    
    OKButtonOldProc = GetWindowLong(OKButtonHwnd, GWL_WNDPROC)
    Call SetWindowLong(OKButtonHwnd, GWL_WNDPROC, GetAddress(AddressOf OKButtonProc))
    
    CancelButtonOldProc = GetWindowLong(CancelButtonHwnd, GWL_WNDPROC)
    Call SetWindowLong(CancelButtonHwnd, GWL_WNDPROC, GetAddress(AddressOf CancelButtonProc))
    
    ComboBoxOldProc = GetWindowLong(ComboBoxHwnd, GWL_WNDPROC)
    Call SetWindowLong(ComboBoxHwnd, GWL_WNDPROC, GetAddress(AddressOf ComboBoxProc))
    
    'dont use following two lines- VB crash - but you can try if you like
    'HelpButtonOldProc = GetWindowLong(HelpButtonHwnd, GWL_WNDPROC)
    'Call SetWindowLong(HelpButtonHwnd, GWL_WNDPROC, GetAddress(AddressOf HelpButtonProc))
    

    'fill the combo box with our desired items
    For i = 0 To UBound(AddItems)
            'get max text width in list portion of combo box
            '& set acordingly CB_SETHORIZONTALEXTENT !
            'as result we have horizontal scroll bar
            'on list portion of combo box
            'only if needed otherwise not
            Dim textsize As POINTAPI
            'device context of combo box
            DC = GetWindowDC(ComboBoxHwnd)
            'measurements of text in pixels
            GetTextExtentPoint32 DC, AddItems(i), Len(AddItems(i)), textsize
            'set CB_SETHORIZONTALEXTENT !
            SendMessage ComboBoxHwnd, CB_SETHORIZONTALEXTENT, textsize.x, ByVal 0&
    'add our items
    a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr(AddItems(i)))
    Next i
    'a = SendMessage(OKButtonHwnd, BM_SETCHECK, 1, 0)
    ' set focus to combobox
    SetFocus ComboBoxHwnd
    
    
    CreateWindows = (FormWindowHwnd <> 0)
    
    
    'find text portion of CB Hwnd
    'TextPortionOfComboBoxHwnd = FindWindowEx(ComboBoxHwnd, ByVal 0&, "EDIT", vbNullString)
    'TextPortionOfComboBoxOldProc = GetWindowLong(TextPortionOfComboBoxHwnd, GWL_WNDPROC)
    'Call SetWindowLong(TextPortionOfComboBoxHwnd, GWL_WNDPROC, GetAddress(AddressOf TextPortionOfComboBoxProc))
    
    
    'auto drop down list
    'b = SendMessage(ComboBoxHwnd, CB_SHOWDROPDOWN, &H0, &H0)
End Function
Public Function WndProc(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  ''This our default window procedure for the window. It will handle all
  ''of our incoming window messages and we will write code based on the
  ''window message what the program should do.

  Dim strTemp As String

    Select Case Message
       Case WM_DESTROY:
          ''Since DefWindowProc doesn't automatically call
          ''PostQuitMessage (WM_QUIT). We need to do it ourselves.
          ''You can use DestroyWindow to get rid of the window manually.
          
          'following two lines for HelpButton
          'Call UnhookWindowsHookEx(WHook)
          
          'next line enables destruction of HelpButton
          Call SetParent(HelpButtonHwnd, FormWindowHwnd)
          
          
          Call PostQuitMessage(0&)
      Case 132 'WM_NCPAINT = &H85
         'ignore this coz it is constantly appearing in case
         'mouse moving on the edge of combo box
      Case 32 ' Const WM_SETCURSOR = &H20
          'ignore this coz it is constantly appearing
      Case 512 ' WM_MOUSEFIRST = &H200
          'ignore this coz it is constantly appearing in case
          'mouse moving on the edge of combo box
      Case 160 ' mouse move in non client area of our form
      
      Case 3, 133 'WM_NCPAINT = &H85 or WM_MOVE = &H3
      'reset our help button position
      Dim FormWindowRect As Rect
      GetWindowRect FormWindowHwnd, FormWindowRect
      SetWindowPos HelpButtonHwnd, 0, FormWindowRect.Right - 40, FormWindowRect.Top + 6, 17, 14, SWP_FRAMECHANGED
       Case Else
       'Debug.Print Message & ", &H" & Hex(Message) & "," & wParam & "," & lParam
    End Select
    

  ''Let windows call the default window procedure since we're done.
  WndProc = DefWindowProc(hwnd&, Message, wParam&, lParam&)

End Function
Public Function OKButtonProc(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case Message
       Case WM_LBUTTONUP:
       exitprocedure
       Case Else
    End Select
    
  ''Since in MyCreateWindow we made the default window proc
  ''this procedure, we have to call the old one using CallWindowProc
  OKButtonProc = CallWindowProc(OKButtonOldProc, hwnd&, Message, wParam&, lParam&)
   
End Function
Public Function CancelButtonProc(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case Message
       Case WM_LBUTTONUP:
       DestroyWindow FormWindowHwnd 'kill our window
    End Select
    
  ''Since in MyCreateWindow we made the default window proc
  ''this procedure, we have to call the old one using CallWindowProc
  CancelButtonProc = CallWindowProc(CancelButtonOldProc, hwnd&, Message, wParam&, lParam&)
   
End Function
Public Function ComboBoxProc(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
     
     
    Select Case Message
 '   Case 307
 '   GoTo 10
    
 '   CharCountNew = SendMessage(ComboBoxHwnd, WM_GETTEXTLENGTH, 0, 0)
    
 '   If CharCountOld >= CharCountNew Then GoTo 10 Else
    
 '  a = SendMessage(ComboBoxHwnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
 '  Dim buffer As String
 '  buffer = String(a, Chr$(0))
 '  b = SendMessage(ComboBoxHwnd, WM_GETTEXT, -1, ByVal buffer)
    'now buffer contains our text in combo box
    
 '    For i = 0 To UBound(AddItems) ' loop thrue array and look at items
 '    s = Left(AddItems(i), Len(buffer))
 '           If Left(AddItems(i), Len(MyStr)) = buffer Then
 '           SetWindowText ComboBoxHwnd, AddItems(i) ' ok set combobox text to found item
 '           a = SendMessage(lParam, EM_SETSEL, LenghtOfCurrentString, Len(AddItems(i)))
            'Beep
            'avoid = True
            
            'DoEvents
            'SendMessage ComboBoxHwnd, WM_UPDATEUISTATE, 0, 0
            'SetFocus ComboBoxHwnd
   '         Exit For
   '         End If
  '  Next i
 '   End If 'od goto 10
'10
    

    'Dim buffer As String
    'buffer = String(GetWindowTextLength(ComboBoxHwnd) + 1, Chr$(0))
    'Get the window's text
    'a = GetWindowText(ComboBoxHwnd, buffer, Len(buffer))
    'Beep
    'Debug.Print buffer
    'Case 308 'WM_CTLCOLORLISTBOX = &H134, some action in list portion
    
    Case 342
        ' appears after keyup and down on list portion
        'emidiatly after that 343 message appears so:
        trouble = True
    Case 343 ' enter was pressed
        If trouble = True Then
        trouble = False
        GoTo 9
        End If
    ComboBoxProc = CallWindowProc(ComboBoxOldProc, hwnd&, Message, wParam&, lParam&)
    'here we use Post message instad Send mesage otherwise everything goes to hell
    PostMessage OKButtonHwnd, BM_CLICK, 0, 0
9
    Case 132 'WM_NCPAINT = &H85
    'ignore this coz it is constantly appearing in case
    'mouse moving on the edge of combo box
    Case 32 ' Const WM_SETCURSOR = &H20
    'ignore this coz it is constantly appearing
    Case 512 ' WM_MOUSEFIRST = &H200
    'ignore this coz it is constantly appearing in case
    'mouse moving on the edge of combo box
    Case 273 ' hex 111 WM_COMMAND, on any key press + last event after selection from list portion of CB
        Select Case wParam
        Case 16778217 ' got focus or something like that
        Case 66536 ' after selection from list portion of combo box
        Case 67109865 'key down I think
            'OK we capture entire text, including last action here
            CharCountNew = SendMessage(ComboBoxHwnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
            Dim buffer As String
            buffer = String(CharCountNew, Chr$(0))
            b = SendMessage(ComboBoxHwnd, WM_GETTEXT, -1, ByVal buffer)
            'now buffer contains our text in combo box
            'Debug.Print "273," & buffer
            
                If CharCountOld < CharCountNew Then '
                'execute autofill text portion
                    'get index of first item starting with "buffer" string
                    a = SendMessage(ComboBoxHwnd, CB_FINDSTRING, -1, ByVal CStr(buffer))
                    'get this item text lenght
                    b = SendMessage(ComboBoxHwnd, CB_GETLBTEXTLEN, a, ByVal 0&)
                        If b = -1 Then 'means there is nothing found in list box port
                        'jump a little to avoid fatal error
                        GoTo 10
                        End If
                    'create buffer1
                    Dim buffer1 As String
                    buffer1 = String(b, Chr$(0))
                    Beep
                    'get this item text into buffer1
                    C = SendMessage(ComboBoxHwnd, CB_GETLBTEXT, a, ByVal buffer1)
                    'set combobox text to our yust found string
                    d = SendMessage(ComboBoxHwnd, WM_SETTEXT, 0, ByVal CStr(buffer1))
                    ' select auto added portion of string, lParam = text box portion hWnd of our combo box
                    e = SendMessage(lParam, EM_SETSEL, CharCountNew, 0)
                    Beep
10
                'and do this
                CharCountOld = CharCountNew
                Else
                'do nothing as back space was pressed
                CharCountOld = CharCountOld - 1
                If CharCountOld < 0 Then CharCountOld = 0
                End If

        Case 50332649 ' key up I think
        Case 83952617 ' limited text - i.e you typed to max extend possible, combo doesn't receive amy more keyboard inputs that is in case there is no WS_AUTOHSCROLL style
        Case Else
        'Debug.Print Message & ", &H" & Hex(Message) & "," & wParam & "," & lParam
        End Select
    Case 307 'generaly something pressed on keyboard
        Select Case wParam
        'Case 1518 ' somewhere at loading
        'Debug.Print Message & ", &H" & Hex(Message) & "," & wParam & "," & lParam
        Case Else
        'Debug.Print Message & ", &H" & Hex(Message) & "," & wParam & "," & lParam
        End Select
    Case Else
    'Debug.Print Message & ", &H" & Hex(Message) & "," & wParam & "," & lParam
    End Select
    
    
    
    
    
    'Case 15
    'ignore this one
    'Case 100
    '    Select Case lParam
    '    Case 17
    '    Case Else
    '    Debug.Print Message & ", &H" & Hex(Message) & "," & wParam & "," & lParam
    '    End Select
        
    'Case 307
    'first time 307 appears during loading
    'lparam of 307 = TextPortionOfComboBoxHwnd
    'we will make function to capture events from TextPortionOfComboBox
    'We need this since inputs from keyboard doesn't go to combo box i.e.
    'combobox doesn't aware which key was pressed
    'We need this avareness to make autotype ability of combobox.
    'If FirstTime307Passed = False Then
    '    FirstTime307Passed = True
    '    TextPortionOfComboBoxHwnd = lParam
    '    TextPortionOfComboBoxOldProc = GetWindowLong(TextPortionOfComboBoxHwnd, GWL_WNDPROC)
    '    Call SetWindowLong(TextPortionOfComboBoxHwnd, GWL_WNDPROC, GetAddress(AddressOf TextPortionOfComboBoxProc))
    'End If
    
    
    'Case 514 ' ignore this
    'Case 1060 ' ignore this
    'Case 7 ' ignore this
    'Case 8 ' ignore this
    'Case 12 ' ignore this &HC WM_SETTEXT
    'Case 13 'WM_GETTEXT = &HD
    
    'CharCountNew = wParam
    '    If CharCountNew < CharCountOld Then
    '    avoid = True
    '    End If
    'CharCountOld = CharCountNew
    'Case 14 'WM_GETTEXTLENGTH = hex &HE
    
    'Case 33 'ignore this
    'l = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal CStr(Combo1.Text))
    'Case 308 'mouse move above certain item
    'a = SendMessage(ComboBoxHwnd, CB_GETCURSEL, 0, 0) 'items index
    'Dim b As String
    'b = SendMessage(ComboBoxHwnd, CB_GETLBTEXTLEN, a, &H0) 'items text lenght
    'SetWindowText TextHwnd, b
    'Case 70 ' like Combo_Change() event in VB
    'Case 32 'mouse over
   
    'Case 307 ' hex 133 WM_CTLCOLOREDIT
    'Beep
    'Case 70 ' hex 46 WM_WINDOWPOSCHANGING
    
    'SetWindowText TextHwnd, Message & "," & wParam & "," & lParam
    'If avoid Then avoid = False: GoTo 10
    
    
    'get current text in text portion of combobox
    'Dim MyStr As String
    'Create a buffer
    'MyStr = String(GetWindowTextLength(ComboBoxHwnd) + 1, Chr$(0))
    'Get the window's text
    'a = GetWindowText(ComboBoxHwnd, MyStr, Len(MyStr))
    'MyStr = Left(MyStr, Len(MyStr) - 1)
    'Dim LenghtOfCurrentString As Integer ' will hold typed in text lenght
    'LenghtOfCurrentString = Len(MyStr)
    
    
    'If MyStr = "" Then GoTo 10 ' not interested for our for ... next
    
   ' For i = 0 To UBound(AddItems) ' loop thrue array and look at items
   ' s = Left(AddItems(i), Len(MyStr))
    '    If Left(AddItems(i), Len(MyStr)) = MyStr Then
    '    SetWindowText ComboBoxHwnd, AddItems(i) ' ok set combobox text to found item
    '    a = SendMessage(lParam, EM_SETSEL, LenghtOfCurrentString, Len(AddItems(i)))
        'Beep
        'avoid = True
        
        'DoEvents
        'SendMessage ComboBoxHwnd, WM_UPDATEUISTATE, 0, 0
        'SetFocus ComboBoxHwnd
        'Exit For
        'End If
    'Next i
    
    
'10
    'get first items index in combo box list portion that has starts same as MyStr
    ' l = -1 if nothing found or items index if found
    'l = SendMessage(ComboBoxHwnd, CB_FINDSTRING, -1, ByVal CStr(MyStr))
    'If l <> -1 Then
    'm = SendMessage(ComboBoxHwnd, CB_GETITEMTEXT, l, 4)
    'End If
    'Beep
    'Beep
    
    'Case 343? hex 157 CB_GETDROPPEDSTATE on any ENTER key on combo box
    'Case Else
    'Beep
     'a = SendMessage(ComboBoxHwnd, CB_ADDSTRING, -1, ByVal CStr(Message))
    'SetWindowText TextHwnd, Message & "," & wParam & "," & lParam
    
     
    'End Select
    
    
    
  ''Since in MyCreateWindow we made the default window proc
  ''this procedure, we have to call the old one using CallWindowProc
  ComboBoxProc = CallWindowProc(ComboBoxOldProc, hwnd&, Message, wParam&, lParam&)
   
End Function



'Public Function TextPortionOfComboBoxProc(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'SendMessage ComboBoxHwnd, Message, wParam, lParam
'Select Case Message
'Case 15
''Case 257
'Case Else
'Debug.Print "aaa" & Message & ", &H" & Hex(Message) & "," & wParam & "," & lParam
'End Select
''Since in MyCreateWindow we made the default window proc
  ''this procedure, we have to call the old one using CallWindowProc
'  TextPortionOfComboBoxProc = CallWindowProc(TextPortionOfComboBoxOldProc, hwnd&, Message, wParam&, lParam&)
'End Function


Public Function HelpButtonProc(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Beep
HelpButtonProc = CallWindowProc(HelpButtonOldProc, hwnd&, Message, wParam&, lParam&)
End Function
Public Function GetAddress(ByVal lngAddr As Long) As Long
    ''Used with AddressOf to return the address in memory of a procedure.

    GetAddress = lngAddr&
    
End Function

Private Sub exitprocedure()

                 Dim MyStr As String
                 'Create a buffer
                 MyStr = String(GetWindowTextLength(ComboBoxHwnd) + 1, Chr$(0))
                 'Get the window's text
                 a = GetWindowText(ComboBoxHwnd, MyStr, Len(MyStr))
                 If Len(MyStr) = 1 Then GoTo 11 ' if user didnt type anything
                 IBCBSelectedItem = MyStr
                 
                 MsgBox "Selection is now waiting as Public IBCBSelectedItem for further use, you selected:" & IBCBSelectedItem
                 
                 'Terminate the window hooking
                 'Call UnhookWindowsHookEx(WHook)
                 'Call SetParent(HelpButtonHwnd, FormWindowHwnd)
11
                 
                 DestroyWindow FormWindowHwnd 'kill our window

End Sub

'Public Function HelpButtonProc(ByVal nCode&, ByVal wParam&, Inf As CWPSTRUCT)
'    Beep
'End Function
