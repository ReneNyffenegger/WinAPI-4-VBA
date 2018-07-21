option explicit

dim hWndListBox as longPtr

function getAddressOfCallback(addr as long) as long
         getAddressOfCallback=addr
end function


function main()

    dim hWnd        as longPtr
    dim regClass    as integer
    dim tq84Class   as WNDCLASSEX
    dim msg_        as MSG
  
    tq84Class.cbSize        = lenB(tq84Class)
    tq84Class.style         = CS_HREDRAW or CS_VREDRAW
    tq84Class.lpfnwndproc   = getAddressOfCallback(AddressOf WindowProc)
    tq84Class.cbClsextra    = 0
    tq84Class.cbWndExtra    = 0
    tq84Class.hInstance     = 0
    tq84Class.hIcon         = LoadIcon      (0, IDI_APPLICATION)
    tq84Class.hIconSm       = LoadIcon      (0, IDI_APPLICATION)
    tq84Class.hCursor       = LoadCursor    (0, IDC_ARROW      )
    tq84Class.hbrBackground = GetStockObject(   WHITE_BRUSH    )
    tq84Class.lpszMenuName  = 0
    tq84Class.lpszClassName ="TQ84CLASS"

    regClass= RegisterClassEx(tq84Class)

    debug.print "creating main window"
    hWnd = CreateWindowEx(                                                _
              0                                                         , _
             "TQ84CLASS"                                                , _
             "Title of window"                                          , _
              WS_OVERLAPPEDWINDOW                                       , _
              CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,  _
              0, 0, 0, 0)

    debug.print "main window created, hWnd = " & hWnd

    if hWnd = 0 then
       msgBox "Failed to create window"
       exit function
    end if

    ShowWindow   hWnd, SW_SHOWNORMAL
    debug.print "main window is shown, hWndListBox = " & hWndListBox

  '
  ' Filling some items into the list box with SendMessage.
  ' Note the »byVal« for the strings.
  '
    SendMessage hWndListBox, LB_ADDSTRING, 0, byVal "Foo"
    SendMessage hWndListBox, LB_ADDSTRING, 0, byVal "Bar"
    SendMessage hWndListBox, LB_ADDSTRING, 0, byVal "Baz"
    dim i as long
    for i = 1 to 50
       SendMessage hWndListBox, LB_ADDSTRING, 0, byVal "i = " & i
    next i


 '  UpdateWindow hWnd
 '  SetFocus     hWnd       
   
    do while 0 <> GetMessage(msg_, 0, 0, 0)
        TranslateMessage msg_
        DispatchMessage  msg_
    loop
   
end function

function WindowProc(              _
         byVal hWnd   as longPtr, _
         byVal msg_   as long   , _
         byVal wParam as long   , _
         byVal lParam as long) as longPtr

    dim ps         as PAINTSTRUCT
    dim clientRect as RECT
    dim hdc        as longPtr

    dim text as string

    select case msg_

      case WM_CREATE

        debug.print "Creating listbox, hWnd = " & hWnd

        hWndListBox = CreateWindowEx(                                 _
              0                                                     , _
             "listbox"                                              , _
              0                                                     , _
              LBS_HASSTRINGS or WS_CHILD or WS_VISIBLE or WS_VSCROLL, _
              0, 0, 0, 0,                                             _
              hWnd                                                  , _
              0, 0, 0 )


        if hWndListBox = 0 then
           msgBox "Could not create listbox, error = " & GetLastError()
           WindowProc = -1
        else
           WindowProc = 0
        end if
        exit function

      case WM_SIZE

        if hWndListBox <> 0 then
           MoveWindow hWndListBox, 0, 0, LOWORD(lParam), HIWORD(lParam), true
        end if
         
      case WM_DESTROY

        PostQuitMessage 0
        exit function

    end select
   
    WindowProc = DefWindowProc(hwnd, msg_, wParam, lParam)

end function
