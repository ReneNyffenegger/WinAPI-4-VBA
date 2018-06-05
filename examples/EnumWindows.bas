option explicit

global r as long

sub main()
    r = 1
    call EnumWindows(addressOf EnumWindowsProc, byVal 0&) 
end sub

function EnumWindowsProc(byVal hWnd as long, byVal lParam as long) as long

    dim windowText as string
    dim retVal     as long
    dim l          as long
    
    windowText = space(GetWindowTextLength(hWnd) + 1)
    retVal     =       GetWindowText(hWnd, windowText, len(windowText))
    windowText = left$(windowText, retVal)
    cells(r, 1) = windowText & " < "
    
    r = r + 1
    
    EnumWindowsProc = true

end function