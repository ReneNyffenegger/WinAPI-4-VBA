option explicit


' { Types
type INPUT_     '   typedef struct tagINPUT ' {
  dwType      as long
  wVK         as integer
  wScan       as integer         '               KEYBDINPUT ki;
  dwFlags     as long            '               HARDWAREINPUT hi;
  dwTime      as long            '           '};
  dwExtraInfo as long            '   '} INPUT, *PINPUT;
  dwPadding   as currency        '   8 extra bytes, because mouses take more.
end type ' }


private type IMAGE_DATA_DIRECTORY ' WinNT.h {
    RVA  as long ' Relative VA
    Size as long
end type ' }

private const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16

type IMAGE_EXPORT_DIRECTORY_TABLE ' WinNT.h {
  Characteristics         as  long
  TimeDateStamp           as  long
  MajorVersion            as  integer
  MinorVersion            as  integer
  Name                    as  long
  Base                    as  long
  NumberOfFunctions       as  long
  NumberOfNames           as  long
  AddressOfFunctions      as  long ' Relative virtual address (RVA) from base of image. Points to an array of RVAs of functions/symbols in the module
  AddressOfNames          as  long ' Relative virtual address (RVA) from base of image
  AddressOfNameOrdinals   as  long ' Relative virtual address (RVA) from base of image
end type ' }

type IMAGE_OPTIONAL_HEADER32 ' WinNT.h {
'
'  2019-01-24: Renamed from IMAGE_OPTIONAL_HEADER
'
 ' Standard fields:
   Magic                                                 as  integer
   MajorLinkerVersion                                    as  byte
   MinorLinkerVersion                                    as  byte
   SizeOfCode                                            as  long
   SizeOfInitializedData                                 as  long
   SizeOfUninitializedData                               as  long
   AddressOfEntryPoint                                   as  long
   BaseOfCode                                            as  long
   BaseOfData                                            as  long

 ' NT additional fields:
   ImageBase                                             as  long
   SectionAlignment                                      as  long
   FileAlignment                                         as  long
   MajorOperatingSystemVersion                           as  integer
   MinorOperatingSystemVersion                           as  integer
   MajorImageVersion                                     as  integer
   MinorImageVersion                                     as  integer
   MajorSubsystemVersion                                 as  integer
   MinorSubsystemVersion                                 as  integer
   Win32VersionValue                                     as  long
   SizeOfImage                                           as  long
   SizeOfHeaders                                         as  long
   CheckSum                                              as  long
   Subsystem                                             as  integer
   DllCharacteristics                                    as  integer
   SizeOfStackReserve                                    as  long
   SizeOfStackCommit                                     as  long
   SizeOfHeapReserve                                     as  long
   SizeOfHeapCommit                                      as  long
   LoaderFlags                                           as  long
   NumberOfRvaAndSizes                                   as  long
'  Data directories
   DataDirectory(0 to IMAGE_NUMBEROF_DIRECTORY_ENTRIES)  as  IMAGE_DATA_DIRECTORY ' 17*8 + 96 = 232
end type ' }

type IMAGE_FILE_HEADER ' { WinNT.h / COFF file header
'
'  2019-01-24: Renamed from IMAGE_COFF_HEADER
'
   Machine               as  integer
   NumberOfSections      as  integer
   TimeDateStamp         as  long
   PointerToSymbolTable  as  long
   NumberOfSymbols       as  long
   SizeOfOptionalHeader  as  integer
   Characteristics       as  integer
end type ' }

type IMAGE_NT_HEADERS32 ' WinNT.h {
  '
  ' 2019-01-24: renamed from  »IMAGE_PE_FILE_HEADER«.
  '
  ' Compare with IMAGE_NT_HEADERS64, also defined in WinNT.h
  '
    Signature      as long
    FileHeader     as IMAGE_FILE_HEADER
    OptionalHeader as IMAGE_OPTIONAL_HEADER32 '  or IMAGE_OPTIONAL_HEADER64
end type ' }

type WNDCLASSEX ' {
    cbSize         as long
    style          as long        ' See CS_* constants
'   lpfnwndproc    as longPtr
    lpfnwndproc    as long
    cbClsextra     as long
    cbWndExtra     as long
    hInstance      as longPtr
    hIcon          as longPtr
    hCursor        as longPtr
    hbrBackground  as longPtr
'   hInstance      as long
'   hIcon          as long
'   hCursor        as long
'   hbrBackground  as long
    lpszMenuName   as string
    lpszClassName  as string
    hIconSm        as longPtr
'   hIconSm        as long
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

type PROCESS_INFORMATION ' { Used for CreateProcess()
    hProcess    as longPtr
    hThread     as longPtr
    dwProcessId as long
    dwThreadId  as long
end type ' }

type RECT ' {
   left     as long
   top      as long
   right    as long
   bottom   as long
end type ' }

type LIST_ENTRY ' { Used by LOADED_IMAGE
  FLink as long
  Blink as long
end type ' }

'
'    LOADED_IMAGE
'      Is defined in both ImageHlp.h and DbgHelp.h
'
type LOADED_IMAGE ' 48 bytes (46 bytes packed ) ' { Used with MapAndLoad
  ModuleName         as long
  hFile              as long
  MappedAddress      as long ' Base address of mapped file
  FileHeader         as long ' Pointer to IMAGE_NT_HEADERS32 ' (Compare with IMAGE_NT_HEADERS64) -- Note: the pointed to IMAGE_NT_HEADERS32 also has a member named FileHeader.
  LastRvaSection     as long ' Pointer to first COFF section header (section table)? 2019-01-24: Renamed from pLstRvaSection 
  NumberOfSections   as long
  Sections           as long ' Pointer to IMAGE_SECTION_HEADER (First COFF section header (section table)??)
  Characteristics    as long ' Image characteristics value
  fSystemImage       as byte ' bool
  fDOSImage          as byte ' bool
'
' At least in C, the compiler pads the following two (new) members
' with the previous two bytes into 4 byte so that in C, adding
' or omitting them should not change anything.
'
' fReadOnly          as byte ' bool
' Version            as byte ' UCHAR
'
' ----------------------------------------------------------
  Links              as LIST_ENTRY ' 2 longs
  SizeOfImage        as long
end type ' }

public type MODULEINFO ' {
   lpBaseOfDll                as long
   SizeOfImage                as long
   EntryPoint                 as long
end type ' }

type MSG ' {
    hWnd    as long
    message as long
    wParam  as long
    lParam  as long
    time    as long
    pt      as POINTAPI
end type ' }

type PAINTSTRUCT ' {
 '  hdc                  as longPtr
    hdc                  as long
    fErase               as long
    rcPaint              as RECT
    fRestore             as long
    fIncUpdate           as long
    rgbReserved(0 to 31) as byte
    'rgbReserved(32) as byte 'this was declared incorrectly in VB API viewer
end type ' }

type SECURITY_ATTRIBUTES ' { Used for CreateProcess()
    nLength              as long
    lpSecurityDescriptor as longPtr
    bInheritHandle       as long
end type ' }

type STARTUPINFO ' { Used for CreateProcess
    cb              as long
    lpReserved      as string
    lpDesktop       as string
    lpTitle         as string
    dwX             as long
    dwY             as long
    dwXSize         as long
    dwYSize         as long
    dwXCountChars   as long
    dwYCountChars   as long
    dwFillAttribute as long
    dwFlags         as long
    wShowWindow     as integer
    cbReserved2     as integer
    lpReserved2     as byte
    hStdInput       as longPtr
    hStdOutput      as longPtr
    hStdError       as longPtr
end type ' }

' }
' { Constants

' { C

' CS_* - WindowClass styles / see WNDCLASSEX {
public const CS_HREDRAW                     = &H2
public const CS_VREDRAW                     = &H1
' }

public const CF_TEXT                        = 1

' Process Creation Flags, used for CreateProcess {

public const CREATE_NEW_CONSOLE      = &h00000010
public const CREATE_SUSPENDED        = &h00000004
public const DEBUG_ONLY_THIS_PROCESS = &h00000002

' }

' Used in CreateWindowEx to indicate default position and dimension.
public const CW_USEDEFAULT  = &H80000000

' }
' { D
' { DT constants, used for DrawText
public const DT_CENTER     = &h01
public const DT_SINGLELINE = &h20
public const DT_VCENTER    = &h04
' }
' }
' { E
' EVENT_* constants ' {
'      Used for SetWinEventHook
public const EVENT_OBJECT_CREATE     = &h8000
public const EVENT_OBJECT_DESTROY    = &h8001
public const EVENT_OBJECT_SHOW       = &h8002
public const EVENT_SYSTEM_FOREGROUND = &h0003
' }
' }
' { F
public const FORMAT_MESSAGE_FROM_SYSTEM    = &h1000
public const FORMAT_MESSAGE_IGNORE_INSERTS = &h0200
public const FORMAT_MESSAGE_TEXT_LEN       = &h00a0
' }
' { G
public const GHND                          = &h42
' }
' { H
public const HC_ACTION               = 0

' Used for nCode in CBTProc (SetWindowsHookEx / WH_CBT).
public const HCBT_ACTIVATE           = 5

public const HSHELL_WINDOWCREATED    = 1  ' Top-level unowned window has been created. Used in WH_SHELL callback.

' { HWND_* contstans
'   These can be used for the parameter hWndInsertAfter in SetWindowPos()
'
public const HWND_BOTTOM    =  1
public const HWND_BROADCAST = &HFFFF&
public const HWND_NOTOPMOST = -2
public const HWND_TOP       =  0
public const HWND_TOPMOST   = -1

' }
' }
' { I
' IDC_ARROW, See -> LoadCursor
public const IDC_ARROW                = 32512&
' IDI_APPLICATION, See -> LoadIcon
public const IDI_APPLICATION          = 32512&

public const IMAGE_DIRECTORY_ENTRY_EXPORT = 0
public const IMAGE_DIRECTORY_ENTRY_IMPORT = 1

' }
' { L

public const LB_ADDSTRING             =  &h0180

' List box styles
public const LBS_HASSTRINGS           =  &H40

public const LOCALE_SNAME             =  &h0000005c

' }
' { M
public const MAX_PATH as integer = 260

' MEM_*: Used for VirtualAlloc or VirtualFree
public const MEM_RESERVE_AND_COMMIT       as long = &H3000
public const MEM_RELEASE                  as long = &H8000

' }
' { P

' Use for flProtect in VirtualAlloc:
public const PAGE_EXECUTE_RW              as long = &H40

public const PM_REMOVE  as long = &H1
' }
' { S

' SM_* constants for GetSystemMetrics() {
public const SM_CMONITORS       = 80 ' Number of display monitors
public const SM_CXSCREEN        =  0 ' Width of primary monitor
public const SM_CYSCREEN        =  1 ' Height of primary monitor
public const SM_SLOWMACHINE     = 73 ' Non zero if slow (low-end) processor
' }

public const SMTO_ABORTIFHUNG        = &H2&

public const SPI_SETNONCLIENTMETRICS = &H2A&

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
public const SW_NORMAL          =  1
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

' }
' { V
public const VK_LBUTTON              = &h001 ' { Virtual keys
public const VK_RBUTTON              = &h002
public const VK_CANCEL               = &h003 ' Implemented as Ctrl-Break on most keyboards
public const VK_MBUTTON              = &h004
public const VK_XBUTTON1             = &h005
public const VK_XBUTTON2             = &h006
public const VK_BACK                 = &h008
public const VK_TAB                  = &h009
public const VK_CLEAR                = &h00c
public const VK_RETURN               = &h00d ' Enter
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
' }
' { W
'  WH_* constants used for SetWindowsHookEx {
'
public const WH_CBT         =  5
public const WH_KEYBOARD_LL = 13 ' Low level keyboard events (compare with WH_KEYBOARD)
public const WH_SHELL       = 10 ' Notification of shell events, such as creation of top level windows.
' }
' { WM_*: Window messsages
public const WM_CHAR          = &h0102
public const WM_CLOSE         = &H0010
public const Wm_CREATE        = &H0001
public const WM_DESTROY       = &H0002
public const WM_KEYDOWN       = &h0100
public const WM_KEYUP         = &h0101
public const WM_PAINT         = &H000F
public const WM_SETTEXT       = &h000C
public const WM_SETTINGCHANGE = &H001A
public const WM_SIZE          = &h0005
public const WM_SYSKEYDOWN    = &h0104
public const WM_SYSKEYUP      = &h0105
' }

' { WS_* / Window styles
public const WS_VISIBLE             = &H10000000
public const WS_VSCROLL             = &H200000
public const WS_TABSTOP             = &H10000
public const WS_THICKFRAME          = &H40000
public const WS_MAXIMIZE            = &H1000000
public const WS_MAXIMIZEBOX         = &H10000
public const WS_MINIMIZE            = &H20000000
public const WS_MINIMIZEBOX         = &H20000
public const WS_SYSMENU             = &H80000
public const WS_BORDER              = &H800000
public const WS_DLGFRAME            = &H400000
public const WS_CAPTION             = WS_BORDER or WS_DLGFRAME
public const WS_CHILD               = &H40000000
public const WS_CHILDWINDOW         = WS_CHILD
public const WS_CLIPCHILDREN        = &H2000000
public const WS_CLIPSIBLINGS        = &H4000000
public const WS_DISABLED            = &H8000000
public const WS_EX_ACCEPTFILES      = &H10&
public const WS_EX_DLGMODALFRAME    = &H1&
public const WS_EX_NOPARENTNOTIFY   = &H4&
public const WS_EX_STATICEDGE       = &H20000
public const WS_EX_TOPMOST          = &H8&
public const WS_EX_TRANSPARENT      = &H20&
public const WS_GROUP               = &H20000
public const WS_HSCROLL             = &H100000
public const WS_ICONIC              = WS_MINIMIZE
public const WS_OVERLAPPED          = &H0&
public const WS_OVERLAPPEDWINDOW    = WS_OVERLAPPED or WS_CAPTION or WS_SYSMENU or WS_THICKFRAME or WS_MINIMIZEBOX or WS_MAXIMIZEBOX
public const WS_POPUP               = &H80000000
public const WS_POPUPWINDOW         = WS_POPUP or WS_BORDER or WS_SYSMENU
public const WS_SIZEBOX             = WS_THICKFRAME
public const WS_TILED               = WS_OVERLAPPED
public const WS_TILEDWINDOW         = WS_OVERLAPPEDWINDOW
' }

public const WINEVENT_OUTOFCONTEXT    = 0
public const WINEVENT_SKIPOWNPROCESS  = 2
' }

' { Brushes - used for GetStockObject
public const WHITE_BRUSH = 0
public const BLACK_BRUSH = 4
' }

' }
#if VBA7 then ' 32-Bit versions of Excel ' {

' { A
    declare function AttachThreadInput lib "user32"                        ( _
       byVal idAttach       as long, _
       byVal idAttachTo     as long, _
       byVal fAttach        as long) as long

' }
' { B
    declare function Beep                lib "kernel32"                    ( _
         byVal dwFreq       as long, _
         byVal dwDuration   as long) as long

    declare ptrSafe function BeginPaint  lib "user32"                      ( _
         byVal hwnd as longPtr, _
         lpPaint    as PAINTSTRUCT) as longPtr

    declare function BringWindowToTop    lib "user32"                      ( _
         byVal lngHWnd      as long) as long

' }
' { C

  ' CallNextHookEx {
    declare function CallNextHookEx      lib "user32"                      ( _
         byVal hHook        as long, _
         byVal nCode        as long, _
         byVal wParam       as long, _
               lParam       as any ) as long
  ' }

    declare function CloseClipboard      lib "User32" () as long

    declare function CloseHandle         lib "kernel32"                    ( _
         byVal hObject            as long) as long

  ' CreateProcess {
    declare ptrSafe function CreateProcess lib "kernel32" alias "CreateProcessA" ( _
         byVal lpApplicationName    as string             , _
         byVal lpCommandLine        as string             , _
         byRef lpProcessAttributes  as SECURITY_ATTRIBUTES, _
         byRef lpThreadAttributes   as SECURITY_ATTRIBUTES, _
         byVal bInheritHandles      as long               , _
         byVal dwCreationFlags      as long               , _
         byRef lpEnvironment        as any                , _
         byVal lpCurrentDirectory   as string             , _
               lpStartupInfo        as STARTUPINFO        , _
               lpProcessInformation as PROCESS_INFORMATION) as longPtr
  ' }

  ' CreateThread {
    declare function CreateThread        lib "kernel32"                    ( _
               lpThreadAttributes as any , _
         byVal dwStackSize        as long, _
         byVal lpStartAddress     as long, _
         byVal lParameter         as long, _
         byVal dwCreationFlags    as long, _
               lpThreadID         as long) as long
  ' }

  ' CreateWindowEx {
  '
  ' See CW_USEDEFAULT for x, y, nWidth and nHeight
  '
    declare ptrSafe function CreateWindowEx lib "user32"      alias "CreateWindowExA" ( _
         byVal dwExStyle       as long   , _
         byVal lpClassName     as string , _
         byVal lpWindowName    as string , _
         byVal dwStyle         as long   , _
         byVal x               as long   , _
         byVal y               as long   , _
         byVal nWidth          as long   , _
         byVal nHeight         as long   , _
         byVal hWndParent      as longPtr, _
         byVal hMenu           as longPtr, _
         byVal hInstance       as longPtr, _
               lpParam as any) as longPtr
    ' }

' }
' { D

  ' DefWindowProc {
    declare ptrSafe function DefWindowProc lib "user32" alias "DefWindowProcA" ( _
         byVal hwnd   as longPtr, _
         byVal wMsg   as long   , _
         byVal wParam as longPtr, _
         byVal lParam as longPtr) as longPtr
  ' }

    declare ptrSafe function DispatchMessage lib "user32" alias "DispatchMessageA" (lpMsg as MSG) as longPtr

  ' { DrawText
  '      Use a combination of DT constants for wFormat
    declare ptrSafe function DrawText     lib "user32" alias "DrawTextA"   ( _
         byVal  hdc     as longPtr, _
         byVal  lpStr   as string , _
         byVal  nCount  as long   , _
                lpRect  as RECT   , _
         byVal  wFormat as long) as long
  ' }
' }
' { E

    declare function EmptyClipboard       lib "User32" () as long

    declare function EnumProcessModules   lib "PSAPI.DLL"                  ( _
         byVal hProcess     as long, _
               lphModule    as long, _
         byVal cb           as long, _
               lpcbNeeded   as long     ) as long

    declare ptrSafe function EndPaint     lib "user32"                     ( _
         byVal hwnd         as longPtr, _
               lpPaint      as PAINTSTRUCT) as long

    declare function EnumChildWindows     lib "user32"                     ( _
         byVal hWndParent   as long, _
         byVal lpEnumFunc   as long, _
         byVal lParam       as long) as long

    declare function EnumWindows         lib "user32"                      ( _
         byVal lpEnumFunc   as long, _
         byVal lParam       as long)   as long

    declare function EnumDisplayMonitors lib "user32"                      ( _
         byVal hdc          as long, _
         byVal lprcClip     as long, _
         byVal lpfnEnum     as long, _
         byVal dwData       as long) as long
' }
' { F

  ' FindWindow: Find the top level window that matches lpClassName (if not null) and lpWindowName (if not null). {
  '
  ' Some Class Names
  '   MS Access:           OMain
  '   MS Excel:            XLMAIN
  '   MS Outlook:          rctrl_renwnd32
  '   MS Word:             OpusApp
  '   Visual Basic Editor: wndclass_desked_gsk
  '
    declare function FindWindow          lib "user32"  alias "FindWindowA" ( _
         byVal lpClassName  as string, _
         byVal lpWindowName as string) as long
  ' }
  ' FindWindowEx {
    declare function FindWindowEx        lib "user32"  alias "FindWindowExA"      ( _
         byVal hWndParent     as long  , _
         byVal hWndChildAfter as long  , _
         byVal lpClassName    as string, _
         byVal lpWindowName   as string) as long
  ' }

    declare ptrSafe function FormatMessage lib "kernel32" alias "FormatMessageA"  ( _
         byVal dwFlags        as long  , _
               lpSource       as any   , _
         byVal dwMessageId    as long  , _
         byVal dwLanguageId   as long  , _
         byVal lpBuffer       as string, _
         byVal nSize          as long  , _
               Arguments      as longPtr) as long
' }
' { G
  ' GetActiveWindow {
  ' Returns the window of the *calling thread* that receives input (Hmmm... is this not the window returned with GetFocus())
  ' Compare GetForegroundWindow
    declare function GetActiveWindow     lib "user32"   () as long
  ' }

  ' GetClassName {
    declare function GetClassName        lib "user32.dll"   alias "GetClassNameA" ( _
         byVal hWnd           as long, _
         byVal lpClassName    as string, _
         byVal nMaxCount      as long) as long
  ' }
  ' GetClientRect {
    declare ptrSafe function GetClientRect lib "user32"                                ( _
         byVal hwnd as longPtr,   _
         lpRect     as RECT  ) as long
  ' }
  ' GetComputerName reads the NetBIOS name from the registry when the system is {
  ' started upc
  '
    declare function GetComputerName     lib "kernel32"     alias "GetComputerNameA"   ( _
         byVal lpBuffer       as string, _
         byRef nSize          as long) as long
  ' }

    declare ptrSafe function GetCurrentProcess lib "kernel32" () as longPtr

    declare function GetCurrentThreadId  lib "kernel32" () as long

    declare function GetCursorPos lib "User32" (lpPoint as POINTAPI) as long

    declare function GetDesktopWindow       lib "user32"   () as long

    declare function GetEnvironmentVariable lib "kernel32" alias "GetEnvironmentVariableA" ( _
         byVal lpName         as string, _
         byVal lpBuffer       as string, _
         byVal nSize          as long) as long

  ' GetForegroundWindow {
  ' Returns the window that currently receives input.
  ' Compare with GetActiveWindow()
    declare function GetForegroundWindow lib "user32"   () as long
  ' }
  ' GetLocaleInfo {
  '      Microsoft says that GetLocaleInfoEx is now preferred (for interoperability reasons).
    declare ptrSafe function GetLocaleInfo lib "kernel32" alias "GetLocaleInfoA"          ( _
         byVal lcid          as long  , _
         byVal dwInfoType    as long  , _
         byVal lpSrcStr      as string, _
         byVal cchSrc        as long  ) as long
  ' }
  ' GetM* {
  ' GetMessage {
  '      When GetMessage (and PeekMessage) encounter a WM_QUIT message, they will
  '      return 0.
    declare ptrSafe function GetMessage  lib "user32" alias "GetMessageA"                 ( _
               lpMsg         as MSG    , _
         byVal hwnd          as longPtr, _
         byVal wMsgFilterMin as long   , _
         byVal wMsgFilterMax as long) as long
  ' }

    declare function GetModuleBaseName    lib "PSAPI.DLL" alias "GetModuleBaseNameA"      ( _
         byVal hProcess   as long  , _
         byVal hModule    as long  , _
         byVal lpFilename as string, _
         byVal nSize      as long) as long

    declare function GetModuleFileNameEx  lib "PSAPI"    alias "GetModuleFileNameExA"     ( _
         byVal hProcess   as long, _
         byVal hModule    as long, _
         byVal lpFilename as string, nSize as long) as boolean

    declare function GetModuleInformation lib "PSAPI"                                     ( _
         byVal hProcess      as long      , _
         byVal hModule       as long      , _
               LPMODULEINFO  as MODULEINFO, _
               cb            as long          ) as boolean
  ' GetM }
  ' GetKeyboardLayout {
    declare function GetKeyboardLayout   lib "user32"       alias "GetKeyboardLayout"     ( _
         byVal idThread      as long )  as long
  ' }

    declare ptrSafe function GetLastError  lib "kernel32" () as long

    declare function GetModuleHandle     lib "kernel32"     alias "GetModuleHandleA"      ( _
         byVal lpModuleName as string) as long

    declare function GetParent           lib "user32"                                     ( _
         byVal hwnd           as long  ) as long

    declare function GetProcAddress      lib "kernel32"                                   ( _
         byVal module         as long  , _
         byVal procName       as string) as long

  ' GetProcessHeap {
  ' - See also HeapAlloc()
    declare ptrSafe function GetProcessHeap lib "kernel32" () As longPtr
  ' }

    ' { GetStockObject
    '   See also predefined brushes (such as WHITE_BRUSH, BLACK_BRUSH etc.)
    declare ptrSafe function GetStockObject lib "gdi32"                                   ( _
         byVal nIndex         as long  ) as longPtr
    ' }
  ' GetSystemDefaultLangID {
    declare ptrSafe function GetSystemDefaultLangID lib "kernel32" alias "GetSystemDefaultLangID" () as integer
  ' }

  ' GetSystemMetric {
  '    nIndex is one of the SM_* constants.
    declare ptrSafe function GetSystemMetrics       lib "user32"                          ( _
         byVal nIndex         as long) as long
  ' }

    declare function GetTempFileName     lib "kernel32"     alias "GetTempFileNameA"      ( _
         byVal lpszPath       as string, _
         byVal lpPrefixString as string, _
         byVal uUnique        as long,   _
         byVal lpTempFileName as string) as long

    declare function GetTempPath         lib "kernel32" alias "GetTempPathA"              ( _
         byVal nBufferLength  as long,  _
         byVal lpBuffer       as string) as long

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

    declare function GlobalAlloc         lib "kernel32"                                  ( _
         byVal wFlags         as long, _
         byVal dwBytes        as long) as long

  ' GlobalLock {
  '      Compare with GlobalUnlock
    declare function GlobalLock          lib "kernel32"                                  ( _
         byVal hMem          as long) as long
  ' }

  ' GlobalUnlock {
  '     Compare with GlobalLock
    declare function GlobalUnlock        lib "kernel32"                                  ( _
         byVal hMem           as long   ) as long
  ' }

' }
' { H

  ' HeapAlloc {
  '
  ' hHeap: Use GetProcessHeap() or HeapCreate
  '
  ' - Compare with -> VirtualAlloc
  '
    declare ptrSafe function HeapAlloc lib "kernel32"                                     ( _
         byVal hHeap          as longPtr, _
         byVal dwFlags        as long   , _
         byVal dwBytes        as longPtr) as longPtr
   ' }

' }
' { I
    declare function IsIconic            lib "user32"                                     ( _
         byVal hwnd           as long) as long

    declare function ImageRvaToVa        lib "Imagehlp.dll"                               ( _
         byVal NTHeaders      as long, _
         byVal Base           as long, _
         byVal RVA            as long, _
         byVal LastRvaSection as long) as long

' }
' { L


    declare function lstrlen             lib "kernel32" alias "lstrlenA"                  ( _
         byVal lpsz      as long) as long

    declare function lstrcpy lib "kernel32"                                               ( _
         byVal lpString1 as any, _
         byVal lpString2 as any) as long

  ' LoadCursor {
  '      See also constants IDC_ARROW etc
    declare ptrSafe function LoadCursor  lib "user32"       alias "LoadCursorA"           ( _
         byVal hInstance      as longPtr, _
         byVal lpCursorName   as string ) as longPtr
  ' }

  ' LoadIcon {
  '   See also -> IDI_APPLICATION
    declare ptrSafe function LoadIcon           lib "user32" alias "LoadIconA"             ( _
         byVal hInstance      as longPtr,   _
         byVal lpIconName     as string) as longPtr
  ' }
  ' LoadKeyboardLayout {
  ' TODO: VkKeyScanEx loadkeyboardlayout makelangid LANG_SYSTEM_DEFAULT LANG_USER_DEFAULT
    declare ptrSafe function LoadKeyboardLayout lib "user32" alias "LoadKeyboardLayoutA"   ( _
         byVal pwszKLID       as string,  _
         byVal flags          as long)   as longPtr
  ' }

' }
' { M

  ' Map* {
    declare function MapAndLoad          lib "Imagehlp.dll"                               ( _
         byVal ImageName      as string,       _
         byVal DLLPath        as string,       _
               LoadedImage    as LOADED_IMAGE, _
               DotDLL         as long,         _
               ReadOnly       as long) as long

    declare function MapVirtualKey       lib "user32"       alias "MapVirtualKeyA"        ( _
         byVal wCode          as long,   _
         byVal wMapType       as long) as long

    declare function MapVirtualKeyEx     lib "user32"       alias "MapVirtualKeyExA"      ( _
         byVal wCode          as long, _
         byVal wMapType       as long, _
         byVal dwhkl          as long) as long

  ' Map }
  
  
    declare ptrSafe function MessageBox  lib "user32"       alias "MessageBoxW"           ( _
         byVal hwnd           as longPtr, _
         byVal lpText         as longPtr, _
         byVal lpCaption      as longPtr, _
         byVal wType          as long) as long

    declare ptrSafe function MoveWindow  lib "user32"       alias "MoveWindow"            ( _
         byVal hwnd           as longPtr, _
         byVal x              as long   , _
         byVal y              as long   , _
         byVal nWidth         as long   , _
         byVal nHeight        as long   , _
         byVal bRepaint       as long ) as long

' }
' { O
    declare function OpenClipboard       lib "User32"                                    ( _
         byVal hwnd           as long) as long


    declare function OpenProcess         lib "kernel32"                                  ( _
         byVal dwDesiredAccess as long, _
         byVal bInheritHandle  as long, _
         byVal dwProcessId as long) as long
' }
' { P

  ' PeekMessage {
  '      See also -> GetMessage
    declare function PeekMessage         lib "user32"       alias "PeekMessageA" ( _
         byRef lpMsg          as MSG , _
         byVal hwnd           as long, _
         byVal wMsgFilterMin  as long, _
         byVal wMsgFilterMax  as long, _
         byVal wRemoveMsg     as long) as long
  ' }

    declare ptrSafe function PostMessage     lib "user32"   alias "PostMessageA" ( _
         byVal hwnd   as longPtr, _
         byVal wMsg   as long   , _
         byVal wParam as longPtr, _
               lParam as any) as long

    declare ptrSafe sub PostQuitMessage      lib "user32" (byVal nExitCode as long)

' }
' { R
    ' RegisterClassEx {
    declare ptrSafe function RegisterClassEx lib "user32"   alias "RegisterClassExA"  ( _
               pcWndClassEx   as WNDCLASSEX ) as integer
    ' }

    declare ptrSafe sub      RtlMoveMemory lib "kernel32"                  ( _
         byRef dest               as any    , _
         byRef source             as any    , _
         byVal size               as longPtr)

    declare ptrSafe sub      RtlZeroMemory lib "kernel32"                  ( _
               dest               as any    , _
         byVal length             as long   )


' }
' { S
    declare function SendInput           lib "user32"                                 ( _
         byVal nInputs as long, _
         byRef pInputs as any , _
         byVal cbSize  as long) as long

    declare function SendMessage         lib "user32"       alias "SendMessageA"      ( _
         byVal hwnd   as long, _
         byVal wMsg   as long, _
         byVal wParam as long, _
               lParam as any) as long

    ' SendMessageW {
    '
    '   The wide-character variant of SendMessage is needed if
    '   a string needs to be passed with lParam. See for example
    '      https://renenyffenegger.ch/notes/Windows/registry/environment-variables
    '
    declare function SendMessageW        lib "user32"       alias "SendMessageW"      ( _
         byVal hwnd   as long, _
         byVal wMsg   as long, _
         byVal wParam as long, _
               lParam as any) as long
    ' }

    declare function SendMessageTimeoutW lib "user32"       alias "SendMessageTimeoutW" ( _
         byVal hwnd     as long, _
         ByVal msg      as long, _
         byVal wParam   as long, _
         byVal lParam   as long, _
         byVal fuFlags  as long, _
         byVal uTimeout as long, _
         lpdwResult     as long) as long

    declare function SetClipboardData    lib "User32"                                 ( _
         byVal wFormat as long, _
         byVal hMem    as long) as long

' { Set *

    declare function SetEnvironmentVariable lib "kernel32" alias "SetEnvironmentVariableA" ( _
         byVal lpName  as string, _
         byVal lpValue as string) as long

    declare ptrSafe function SetFocus    lib "user32" (byVal hwnd as longPtr) as longPtr

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

    ' SetWinEventHook. The counter-function is UnhookWinEvent {
    '
    declare function SetWinEventHook     lib "user32.dll"                             ( _
         byVal eventMin          as long  , _
         byVal eventMax          as long  , _
         byVal hmodWinEventProc  as long  , _
         byVal pfnWinEventProc   as long  , _
         byVal idProcess         as long  , _
         byVal idThread          as long  , _
         byVal dwFlags           as long) as long
    ' }
' }
    declare function ShellExecute lib "shell32.dll"         alias "ShellExecuteA"     ( _
         byVal hwnd         as long  , _
         byVal lpOperation  as string, _
         byVal lpFile       as string, _
         byVal lpParameters as string, _
         byVal lpDirectory  as string, _
         byVal lpShowCmd    as long)      as long

  ' ShowWindow {
  '
  ' Use one of the SW*_ constants for nCmdSHow
  '
    declare ptrSafe function ShowWindow          lib "user32" ( _
         byVal hwnd       as long, _
         byVal nCmdSHow   as long) as long
  ' }

    declare sub      Sleep               lib "kernel32" (byVal dwMilliseconds as long   )
' }
' { T
    ' TranslateMessage translates virtual-key messages into character messages. {
    declare ptrSafe function TranslateMessage    lib "user32" (byRef lpMsg as MSG) as long
    ' }
' }
' { U
    ' UnhookWinEvent: unhook hooks established with SetWinEventHook {
    '
    declare function UnhookWinEvent      lib "user32.dll"                             ( _
         byRef hWinEventHook as long) as long
    ' }

    declare         function UnMapAndLoad         lib "Imagehlp.dll" (LoadedImage as LOADED_IMAGE) as long

    declare         function UnhookWindowsHookEx  lib "user32" (byVal hHook as long   ) as long

    declare ptrSafe function UpdateWindow         lib "user32" (byVal hwnd  as longPtr) as long

' }
' { V

  ' VirtualAlloc {
  ' - Use one of the MEM_* constants for flAllocationType
  ' - Use PAGE_* for flProtect
  ' - See also VirtualFree
  ' - Compare with HeapAlloc
    declare ptrSafe function VirtualAlloc lib "kernel32"                              ( _
         byVal address          as longPtr, _
         byVal size             as longPtr, _
         byVal flAllocationType as long   , _
         byVal flProtect        as long) as longPtr
   ' }
   ' VirtualFree {
   ' - See also MEM_*
   ' - See also VirtualAlloc
     declare ptrSafe function VirtualFree lib "kernel32" ( _
         byVal lpAddress      as longPtr, _
         byVal size           as longPtr, _
         byVal dwFreeType     as long   ) as long

   ' }
   ' VirtualProtect {
    declare function VirtualProtect lib "kernel32"                                   ( _
               lpAddress      as long    , _
         byVal dwSize         as long    , _
         byVal flNewProtect   as long    , _
               lpflOldProtect as long) as long
  ' }
    declare function VkKeyScan           lib "user32"       alias "VkKeyScanA"        ( _
         byVal cChar      as byte) as integer

  ' VkKeyScanEx {
  '      Convert a "character" into a *xvirtual key code* and a *shift state*.
  '      dwhkl apparently refers to a kayboard layout which can be optained
  '      with GetKeyboardLayout.
    declare ptrSafe function VkKeyScanEx lib "user32"       alias "VkKeyScanExA"      ( _
         byVal ch         as byte    , _
         byVal dwhkl      as longPtr ) as integer
  ' }

' }
' { W
    declare function WaitMessage         lib "user32" () as long
' }
' }
#else ' 64-Bit versions of Excel ' {

    declare ptrSafe sub Sleep lib "kernel32" (byVal dwMilliseconds as longPtr)


#end if ' }
' Helpers {

function LOBYTE(dw as long) as byte ' {
    LOBYTE = dw and &h000000ff
end function ' }

function LOWORD(dw as long) as long ' {
    if dw And &H8000& then
        LOWORD = &H8000 Or (dw And &H7FFF&)
    else
        LOWORD = dw And &HFFFF&
    end if
end function ' }

function HIWORD(dw as long) as long ' {

    if dw and &H80000000 then
       HIWORD = (dw \ 65535) - 1
    else
       HIWORD = dw \ 65535
    end if

end function ' }

' }
