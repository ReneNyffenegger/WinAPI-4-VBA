option explicit

type INPUT_     '   typedef struct tagINPUT {
  dwType      as long
  wVK         as integer
  wScan       as integer         '               KEYBDINPUT ki;
  dwFlags     as long            '               HARDWAREINPUT hi;
  dwTime      as long            '           };
  dwExtraInfo as long            '   } INPUT, *PINPUT;
  dwPadding   as currency        '   8 extra bytes, because mouses take more.
end type ' }

public const INPUT_KEYBOARD   as long = 1
public const MAPVK_VK_TO_CHAR as long = 2 ' Used in MapVirtualKey
public const MAPVK_VK_TO_VSC  as long = 0 ' Used in MapVirtualKey
public const KEYEVENTF_KEYUP  as long = 2 ' Used for dwFlags in INPUT_


type KBDLLHOOKSTRUCT ' {
     vkCode      as long ' virtual key code in range 1 .. 254
     scanCode    as long ' hardware code
     flags       as long ' bit 4: alt key was pressed
     time        as long
     dwExtraInfo as long
end  type ' }

type POINTAPI ' {
    x as long
    y as long
end type ' }

type RECT ' {
   left     as long
   top      as long
   right    as long
   bottom   as long
end type ' }


type MSG ' {
    hWnd    as long
    message as long
    wParam  as long
    lParam  as long
    time    as long
    pt      as POINTAPI
end type ' }

public const HC_ACTION               = 0

public const HSHELL_WINDOWCREATED    = 1  ' Top-level unowned window has been created. Used in WH_SHELL callback.

' { HWND_* contstans
'   These can be used for the parameter hWndInsertAfter in SetWindowPos()
'
public const HWND_BOTTOM    =  1
public const HWND_NOTOPMOST = -2
public const HWND_TOP       =  0
public const HWND_TOPMOST   = -1

' }

public const PM_REMOVE  as long = &H1

' SW_* constants for ShowWindow() {
public const SW_FORCEMINIMIZE   = 11 ' Minimizes a window.
public const SW_HIDE            =  0 ' Hides the window and activates another window.
public const SW_MAXIMIZE        =  3 ' Maximizes a window.
public const SW_MINIMIZE        =  6 ' Minimizes the specified window and activates the next top-level window.
public const SW_RESTORE         =  9 ' Activates and displays the window.
public const SW_SHOW            =  5 ' Activates the window.
public const SW_SHOWMAXIMIZED   =  3 ' Activates the window and displays it as a maximized window.
public const SW_SHOWMINIMIZED   =  2 ' Activates the window and displays it as a minimized window.
public const SW_SHOWMINNOACTIVE =  7 ' Displays the window as a minimized window (without activating the window).
public const SW_SHOWNA          =  8 ' Displays the window in its current size and position (without activating the window).
public const SW_SHOWNOACTIVATE  =  4 ' Displays a window in its most recent size and position (without activating the window).
public const SW_SHOWNORMAL      =  1 ' Activates and displays a window.
' }

' { SWP_* constants for SetWindowPos
public const SWP_NOSIZE         as long = &h1
public const SWP_NOMOVE         as long = &h2
public const SWP_NOZORDER       as long = &h4
public const SWP_NOREDRAW       as long = &h8
public const SWP_NOACTIVATE     as long = &h10
public const SWP_FRAMECHANGED   as long = &h20
public const SWP_SHOWWINDOW     as long = &h40
public const SWP_HIDEWINDOW     as long = &h80
public const SWP_NOCOPYBITS     as long = &h100
public const SWP_NOOWNERZORDER  as long = &h200
public const SWP_NOSENDCHANGING as long = &h400
public const SWP_DEFERERASE     as long = &h2000
public const SWP_ASYNCWINDOWPOS as long = &h4000
' }

public const VK_LBUTTON              = &h001 ' { Virtual keys
public const VK_RBUTTON              = &h002
public const VK_CANCEL               = &h003 ' Implemented as Ctrl-Break on most keyboards
public const VK_MBUTTON              = &h004
public const VK_XBUTTON1             = &h005
public const VK_XBUTTON2             = &h006
public const VK_BACK                 = &h008
public const VK_TAB                  = &h009
public const VK_CLEAR                = &h00c
public const VK_RETURN               = &h00d
public const VK_SHIFT                = &h010
public const VK_CONTROL              = &h011
public const VK_MENU                 = &h012
public const VK_PAUSE                = &h013
public const VK_CAPITAL              = &h014
public const VK_KANA                 = &h015
public const VK_HANGUEL              = &h015
public const VK_HANGUL               = &h015
public const VK_JUNJA                = &h017
public const VK_FINAL                = &h018
public const VK_HANJA                = &h019
public const VK_KANJI                = &h019
public const VK_ESCAPE               = &h01b
public const VK_CONVERT              = &h01c
public const VK_NONCONVERT           = &h01d
public const VK_ACCEPT               = &h01e
public const VK_MODECHANGE           = &h01f
public const VK_SPACE                = &h020
public const VK_PRIOR                = &h021
public const VK_NEXT                 = &h022
public const VK_END                  = &h023
public const VK_HOME                 = &h024
public const VK_LEFT                 = &h025
public const VK_UP                   = &h026
public const VK_RIGHT                = &h027
public const VK_DOWN                 = &h028
public const VK_SELECT               = &h029
public const VK_PRINT                = &h02a
public const VK_EXECUTE              = &h02b
public const VK_SNAPSHOT             = &h02c
public const VK_INSERT               = &h02d
public const VK_DELETE               = &h02e
public const VK_HELP                 = &h02f
public const VK_LWIN                 = &h05b
public const VK_RWIN                 = &h05c
public const VK_APPS                 = &h05d
public const VK_SLEEP                = &h05f
public const VK_NUMPAD0              = &h060
public const VK_NUMPAD1              = &h061
public const VK_NUMPAD2              = &h062
public const VK_NUMPAD3              = &h063
public const VK_NUMPAD4              = &h064
public const VK_NUMPAD5              = &h065
public const VK_NUMPAD6              = &h066
public const VK_NUMPAD7              = &h067
public const VK_NUMPAD8              = &h068
public const VK_NUMPAD9              = &h069
public const VK_MULTIPLY             = &h06a
public const VK_ADD                  = &h06b
public const VK_SEPARATOR            = &h06c
public const VK_SUBTRACT             = &h06d
public const VK_DECIMAL              = &h06e
public const VK_DIVIDE               = &h06f
public const VK_F1                   = &h070
public const VK_F2                   = &h071
public const VK_F3                   = &h072
public const VK_F4                   = &h073
public const VK_F5                   = &h074
public const VK_F6                   = &h075
public const VK_F7                   = &h076
public const VK_F8                   = &h077
public const VK_F9                   = &h078
public const VK_F10                  = &h079
public const VK_F11                  = &h07a
public const VK_F12                  = &h07b
public const VK_F13                  = &h07c
public const VK_F14                  = &h07d
public const VK_F15                  = &h07e
public const VK_F16                  = &h07f
public const VK_F17                  = &h080
public const VK_F18                  = &h081
public const VK_F19                  = &h082
public const VK_F20                  = &h083
public const VK_F21                  = &h084
public const VK_F22                  = &h085
public const VK_F23                  = &h086
public const VK_F24                  = &h087
public const VK_NUMLOCK              = &h090
public const VK_SCROLL               = &h091
public const VK_LSHIFT               = &h0a0
public const VK_RSHIFT               = &h0a1
public const VK_LCONTROL             = &h0a2
public const VK_RCONTROL             = &h0a3
public const VK_LMENU                = &h0a4 ' This is apparently the left  "Alt" key
public const VK_RMENU                = &h0a5 ' This is apparently the right "Alt" key
public const VK_BROWSER_BACK         = &h0a6
public const VK_BROWSER_FORWARD      = &h0a7
public const VK_BROWSER_REFRESH      = &h0a8
public const VK_BROWSER_STOP         = &h0a9
public const VK_BROWSER_SEARCH       = &h0aa
public const VK_BROWSER_FAVORITES    = &h0ab
public const VK_BROWSER_HOME         = &h0ac
public const VK_VOLUME_MUTE          = &h0ad
public const VK_VOLUME_DOWN          = &h0ae
public const VK_VOLUME_UP            = &h0af
public const VK_MEDIA_NEXT_TRACK     = &h0b0
public const VK_MEDIA_PREV_TRACK     = &h0b1
public const VK_MEDIA_STOP           = &h0b2
public const VK_MEDIA_PLAY_PAUSE     = &h0b3
public const VK_LAUNCH_MAIL          = &h0b4
public const VK_LAUNCH_MEDIA_SELECT  = &h0b5
public const VK_LAUNCH_APP1          = &h0b6
public const VK_LAUNCH_APP2          = &h0b7
public const VK_OEM_1                = &h0ba
public const VK_OEM_PLUS             = &h0bb
public const VK_OEM_COMMA            = &h0bc
public const VK_OEM_MINUS            = &h0bd
public const VK_OEM_PERIOD           = &h0be
public const VK_OEM_2                = &h0bf
public const VK_OEM_3                = &h0c0
public const VK_OEM_4                = &h0db
public const VK_OEM_5                = &h0dc
public const VK_OEM_6                = &h0dd
public const VK_OEM_7                = &h0de
public const VK_OEM_8                = &h0df
public const VK_OEM_102              = &h0e2
public const VK_PROCESSKEY           = &h0e5
public const VK_PACKET               = &h0e7
public const VK_ATTN                 = &h0f6
public const VK_EXSEL                = &h0f8
public const VK_PLAY                 = &h0fa
public const VK_NONAME               = &h0fc ' }

'
'      Constants used for SetWindowsHookEx
'
public const WH_SHELL       = 10 ' Notification of shell events, such as creation of top level windows.
public const WH_KEYBOARD_LL = 13 ' Low level keyboard events

public const WM_CHAR        = &h0102
public const WM_KEYDOWN     = &h0100
public const WM_KEYUP       = &h0101
public const WM_SETTEXT     = &h000C
public const WM_SYSKEYDOWN  = &h0104
public const WM_SYSKEYUP    = &h0105



#if VBA7 then ' 32-Bit versions of Excel ' {

    declare function AttachThreadInput Lib "user32"                        ( _
    	 byVal idAttach       as long, _
    	 byVal idAttachTo     as long, _
    	 byVal fAttach        as long) as long

    declare function Beep                lib "kernel32"                    ( _
         byVal dwFreq       as long, _
         byVal dwDuration   as long) as long

    declare function BringWindowToTop    lib "user32"                      ( _
         byVal lngHWnd      as long) as long

    declare function CallNextHookEx      lib "user32"                      ( _
         byVal hHook        as long, _
         byVal nCode        as long, _
         byVal wParam       as long, _
               lParam       as any ) as long

    declare function EnumChildWindows     lib "user32"                     ( _
         byVal hWndParent   as long, _
         byVal lpEnumFunc   as long, _
         byVal lParam       as long) as long

    declare function EnumWindows         lib "user32"                      ( _
         byVal lpEnumFunc   as long, _
         byVal lParam       as long)   as long

    declare function FindWindow          lib "user32"  alias "FindWindowA" ( _
         byVal lpClassName  as string, _
         byVal lpWindowName as string) as long
  '
  ' Some Class Names
  '   MS Access:           OMain
  '   MS Excel:            XLMAIN
  '   MS Outlook:          rctrl_renwnd32
  '   MS Word:             OpusApp
  '   Visual Basic Editor: wndclass_desked_gsk

    declare function FindWindowEx        lib "user32"  alias "FindWindowExA"      ( _
         byVal hWnd           as long  , _
         byVal hWndChildAfter as long  , _
         byVal lpClassName    as string, _
         byVal lpWindowName   as string) as long

    declare function GetActiveWindow     lib "user32"   () as long

    declare function GetClassName        lib "user32.dll"   alias "GetClassNameA" ( _
         byVal hWnd           as long, _
         byVal lpClassName    as string, _
         byVal nMaxCount      as long) as long

    declare function GetComputerName     lib "kernel32"     alias "GetComputerNameA"   ( _
         byVal lpBuffer       as string, _
               nSize          as long) as long

    declare function GetCurrentThreadId  lib "kernel32" () as long

    declare function GetDesktopWindow    lib "user32"   () as long

    declare function GetForegroundWindow lib "user32"   () as long

    declare function GetKeyboardLayout   lib "user32"       alias "GetKeyboardLayout"     ( _
         byVal pwszKLID       as string) as long

    declare function GetLastError        lib "kernel32" () as long

    declare function GetModuleHandle     lib "kernel32"     alias "GetModuleHandleA"      ( _
         byVal lpModuleName as string) as long

    declare function GetParent           lib "user32"                                     ( _
         byVal hwnd           as long  ) as long

    declare function GetTempFileName     lib "kernel32"     alias "GetTempFileNameA"      ( _
         byVal lpszPath       as string, _
         byVal lpPrefixString as string, _
         byVal uUnique        as long,   _
         byVal lpTempFileName as string) as long

    declare function GetTempPath         lib "kernel32" alias "GetTempPathA"              ( _
         byVal nBufferLength  as long,  _
         byVal lpBuffer       as String) as long

    declare function GetWindowRect       lib "user32.dll"                                 ( _
         byVal hwnd           as long, _
         byRef lpRect as RECT        )   as long

    declare function GetWindowThreadProcessId lib "user32"                                ( _
         byVal hwnd           as long, _
    	         lpdwProcessId  as long) as long

    declare function GetWindowText       lib "user32"       alias "GetWindowTextA"        ( _
         byVal hWnd           as long  , _
         byVal lpString       as string, _
         byVal nMaxCount      as long ) as long

    declare function GetWindowTextLength lib "user32"       alias "GetWindowTextLengthA"  ( _
         byVal hWnd           as long    ) as long

    declare function GetUserName         lib "advapi32.dll" alias "GetUserNameA"          ( _
         byVal lpBuffer       as string, _
               nSize          as long    ) as long

    declare function IsIconic            lib "user32"                                     ( _
         byVal hwnd           as long) as long

    declare function MapVirtualKey       lib "user32"       alias "MapVirtualKeyA"        ( _
         byVal wCode          as long,   _
         byVal wMapType       as long) as long

    declare function MapVirtualKeyEx     lib "user32"       alias "MapVirtualKeyExA"      ( _
         byVal wCode          as long, _
         byVal wMapType       as long, _
         byVal dwhkl          as long) as long

    declare function PeekMessage         lib "user32"       alias "PeekMessageA" ( _
         byRef lpMsg          as MSG , _
         byVal hwnd           as long, _
         byVal wMsgFilterMin  as long, _
         byVal wMsgFilterMax  as long, _
         byVal wRemoveMsg     as long) as long

    declare function PostMessage         lib "user32"       alias "PostMessageA" ( _
         byVal hwnd   as long, _
         byVal wMsg   as long, _
         byVal wParam as long, _
               lParam as any) as long

    declare function SendInput           lib "user32"                                 ( _
         byVal nInputs as long, _
         byRef pInputs as any , _
         byVal cbSize  as long) as long

    declare function SendMessage         lib "user32"       alias "SendMessageA" ( _
         byVal hwnd   as long, _
         byVal wMsg   as long, _
         byVal wParam as long, _
               lParam as any) as long

    declare function SetForegroundWindow lib "user32" (byVal hWnd as long) as long

    declare function SetWindowsHookEx    lib "user32"       alias "SetWindowsHookExA" ( _
         byVal idHook     as long, _
         byVal lpfn       as long, _
         byVal hmod       as long, _
         byVal dwThreadId as long) as long

    declare function SetWindowPos        lib "user32"                                 ( _
         byVal hWnd            as long, _
         byVal hWndInsertAfter as long, _
         byVal x               as long, _
         byVal y               as long, _
         byVal cx              as long, _
         byVal cy              as long, _
         byVal uFlags          as long)   as long
    '
    ' Use one of the SWP_* constants for wFlags
    '

    declare function SetWindowText       lib "user32"       alias "SetWindowTextA"    ( _
         byVal hWnd            as long  , _
         byVal lpString        as string) as long

    declare function ShellExecute Lib "shell32.dll"         alias "ShellExecuteA"     ( _
         byVal hwnd         as long  , _
         byVal lpOperation  as string, _
         byVal lpFile       as string, _
         byVal lpParameters as string, _
         byVal lpDirectory  as string, _
         byval lpShowCmd    as long)      as long

    declare function ShowWindow          lib "user32" ( _
         byVal hwnd       as long, _
         byVal nCmdSHow   as long) as long
  '
  ' Use one of the SW*_ constants for nCmdSHow
  '

    declare sub      Sleep               lib "kernel32" (byVal dwMilliseconds as long   )
    declare function TranslateMessage    lib "user32" (byRef lpMsg as MSG) as long

    declare function UnhookWindowsHookEx lib "user32" (ByVal hHook as long) as long

    declare function VkKeyScan           lib "user32"       alias "VkKeyScanA"        ( _
         byVal cChar      as byte) as integer

    declare function WaitMessage         lib "user32" () as long
' }
#else ' 64-Bit versions of Excel ' {

    declare ptrSafe sub Sleep lib "kernel32" (byVal dwMilliseconds as longPtr)


#end if ' }
