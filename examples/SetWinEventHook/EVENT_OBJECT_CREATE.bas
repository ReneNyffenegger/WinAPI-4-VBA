option explicit

dim hh as long

function getAddressOfCallback(addr as long) as long
  '
  ' https://renenyffenegger.ch/notes/development/languages/VBA/language/operators/addressOf
  '
    getAddressOfCallback=addr
end function

sub main ' {

    dim hWnd        as longPtr
    dim regClass    as integer
    dim tq84Class   as WNDCLASSEX
    dim msg_        as MSG

    tq84Class.cbSize        = lenB(tq84Class)
    tq84Class.style         = CS_HREDRAW or CS_VREDRAW
    tq84Class.lpfnwndproc   = getAddressOfCallback(addressOf WindowProc)
    tq84Class.cbClsextra    = 0
    tq84Class.cbWndExtra    = 0
    tq84Class.hInstance     = application.hInstance
    tq84Class.hIcon         = LoadIcon      (0, IDI_APPLICATION)
    tq84Class.hIconSm       = LoadIcon      (0, IDI_APPLICATION)
    tq84Class.hCursor       = LoadCursor    (0, IDC_ARROW      )
    tq84Class.hbrBackground = GetStockObject(   WHITE_BRUSH    )
    tq84Class.lpszMenuName  = 0
    tq84Class.lpszClassName ="TQ84CLASS"

    regClass= RegisterClassEx(tq84Class)

    hWnd = CreateWindowEx(                  _
              WS_EX_DLGMODALFRAME         , _
             "TQ84CLASS"                  , _
             "Title of window"            , _
              WS_POPUPWINDOW or WS_CAPTION, _
              100, 100, 500, 200,           _
              0, 0,                         _
              application.hInstance       , _
              0)

    if hWnd = 0 then
       msgBox "Failed to create window"
       exit sub
    end if

    ShowWindow   hWnd, SW_SHOWNORMAL

  '
  ' Enter message loop until WM_QUIT is encountered:
  '
    do while 0 <> GetMessage(msg_, 0, 0, 0)
        TranslateMessage msg_
        DispatchMessage  msg_
    loop

end sub ' }

' { WindowProc
function WindowProc(              _
         byVal lhwnd  as longPtr, _
         byVal msg_   as long   , _
         byVal wParam as long   , _
         byVal lParam as long) as longPtr

    dim ps         as PAINTSTRUCT
    dim clientRect as RECT
    dim hdc        as longPtr

    dim text as string

    select case msg_

          case WM_CREATE

              hh = SetWinEventHook(       _
                EVENT_OBJECT_CREATE     , _
                EVENT_OBJECT_CREATE     , _
                0                       , _
                addressOf WinEventProc  , _
                0&                      , _
                0&                      , _
                WINEVENT_OUTOFCONTEXT or WINEVENT_SKIPOWNPROCESS)

              debug.print "hh = " & hh ", last error = " & GetLastError()

          case WM_PAINT

              hdc = BeginPaint(lhwnd, ps)
              call GetClientRect(lhwnd, clientRect)
              text = "SetWinEventHook"

              call DrawText(      _
                     hdc,         _
                     text,        _
                     len(text),   _
                     clientRect,  _
                     DT_SINGLELINE or DT_CENTER or DT_VCENTER)

              call EndPaint(lhwnd, ps)

              exit function

          case WM_DESTROY

              PostQuitMessage 0
              exit function

    end select

    WindowProc = DefWindowProc(lhwnd, msg_, wParam, lParam)

end function ' }

' { WinEventProc
function WinEventProc(               _
      byVal hookHandle     as long,  _
      byVal lEvent         as long,  _
      byVal hWnd           as long,  _
      byVal idObject       as long,  _
      byVal idChild        as long,  _
      byVal idEventThread  as long,  _
      byVal dwmsEventTime  as long   _
  ) as long


  if     lEvent = EVENT_OBJECT_CREATE then
         debug.print "An object was created"
  elseif lEvent = EVENT_OBJECT_DESTROY then
         debug.print "An object was destroyed"
  elseif lEvent = EVENT_SYSTEM_FOREGROUND then
         debug.print "EVENT_SYSTEM_FOREGROUND"
  elseif lEvent = EVENT_OBJECT_SHOW then
         debug.print "EVENT_OBJECT_SHOW"
  else
         debug.print "Unexpected event"
  end if

end function ' }
