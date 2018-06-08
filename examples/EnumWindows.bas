option explicit

global r as long

sub main()
    r = 1
    call EnumWindows(addressOf EnumWindowsProc, byVal 0&) 
end sub

function EnumWindowsProc(byVal hWnd as long, byVal lParam as long) as long ' {

    dim windowText  as string
    dim windowClass as string * 256
    dim retVal      as long
    dim l           as long

    cells(r, 1) = hWnd
    cells(r, 2) = GetParent(hWnd)
    
    windowText = space(GetWindowTextLength(hWnd) + 1)
    retVal     =       GetWindowText(hWnd, windowText, len(windowText))
    windowText = left$(windowText, retVal)
    cells(r, 3) = windowText

    retVal = GetClassName(hWnd, windowClass, 255)
    windowClass = left$(windowClass, retVal)
    cells(r, 4) = windowClass


    
    r = r + 1
    
  '
  ' Return true to indicate that we want to continue
  ' with the enumeration of the windows:
  '
    EnumWindowsProc = true

end function ' }
