option explicit

dim hookHandle  as long 
dim hookStarted as boolean

public cnt as long

function LowLevelKeyboardProc(byVal nCode as long, byVal wParam as long, lParam as KBDLLHOOKSTRUCT) as long ' {

    dim upOrDown as string
    dim altKey   as boolean
    dim char     as string

    if nCode <> HC_ACTION then
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if

    select case wParam
           case WM_KEYDOWN   : upOrDown = "keyDown"
           case WM_KEYUP     : upOrDown = "keyUp"
           case WM_SYSKEYDOWN: upOrDown = "sysDown"
           case WM_SYSKEYUP  : upOrDown = "sysUp"
    end select

  '
  ' Apparently, the 5th bit is set if an ALT key was involved:
  ' 
    altKey = lParam.flags and 32

    if ( lParam.vkCode >= asc("A") ) and ( lParam.vkCode <= asc("Z") ) then
       char = chr(lParam.vkCode)

    else

      select case lParam.vkCode
             case VK_ESCAPE   : char = "esc"

             case VK_LCONTROL : char = "ctrl L"
             case VK_RCONTROL : char = "ctrl R"

             case VK_LMENU    : char = "menu L"
             case VK_RMENU    : char = "menu R"

             case VK_RIGHT    : char = ">>"
             case VK_LEFT     : char = "<<"
             case VK_UP       : char = "^^"
             case VK_DOWN     : char = "vv"

             case VK_LSHIFT   : char = "shift L"
             case VK_RSHIFT   : char = "shift R"

             case VK_LWIN     : char = "win L"
             case VK_RWIN     : char = "win R"

             case VK_RETURN   : char = "enter"
             case else        : char = "?"
      end select

    end if

  '
  ' Display what the user has pressed.
  '(Needs Excel)
  '
    cells(1,1) = upOrDown
    cells(1,2) = lParam.vkCode
    cells(1,3) = char
    cells(1,4) = lParam.flags

    if altKey then cells(1,5) = "alt" else cells(1,5) = "-"

    if lParam.vkCode = VK_ESCAPE then stopHook

    LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)

end function ' }

public sub stopHook() ' {

    if hookStarted then
       UnhookWindowsHookEx hookHandle
       hookStarted = false
    end if

    cells.clear

end sub ' }

sub installHook() ' {

  ' don't hook the keyboard twice !!


    if hookStarted = false then
        hookHandle = SetWindowsHookEx(  _ 
           WH_KEYBOARD_LL                  , _
           addressOf LowLevelKeyboardProc  , _
           application.hInstance           , _
           0 )

        if hookHandle <> 0 then
           hookStarted = true
        else
           msgBox "Could not install hook"
        end if

    else
        msgBox "Hook is already enabled"
    end if

end sub  ' }
