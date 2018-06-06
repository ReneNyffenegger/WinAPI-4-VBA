option explicit

global row_    as long
global indent_ as long

sub main()
    row_    = 1
    indent_ = 1
    call EnumChildWindows(0, addressOf CallBackFunc, byVal 0&) 
end sub

function CallBackFunc(byVal hWnd as long, byVal lParam as long) as long

    dim windowText  as string
    dim windowClass as string * 256
    dim retVal      as long
    dim l           as long

    dim hWndParent  as long

    cells(row_, 1) = hWnd

    hWndParent     = GetParent(hWnd)
    cells(row_, 2) = hWndParent
    

    retVal = GetClassName(hWnd, windowClass, 255)
    windowClass = left$(windowClass, retVal)
    cells(row_, 3) = windowClass

    windowText = space(GetWindowTextLength(hWnd) + 1)
    retVal     =       GetWindowText(hWnd, windowText, len(windowText))
    windowText = left$(windowText, retVal)
    cells(row_, 3 + indent_) = windowText

    
    row_ = row_ + 1

    indent_ = indent_ + 1
'   call EnumChildWindows(hWndParent, addressOf CallBackFunc, byVal 0&)
    
  '
  ' Return true to indicate that we want to continue
  ' with the enumeration of the windows:
  '
    CallBackFunc = true

end function

