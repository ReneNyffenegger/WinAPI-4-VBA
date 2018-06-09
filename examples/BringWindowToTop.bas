option explicit

global captionPart_ as string

sub main(captionPart as string)

    captionPart_ = captionPart

    call EnumWindows(addressOf EnumWindowsProc, byVal 0&) 

  '
  ' If called from the visual basic editor, sleep for a second
  ' so as to see the effect of the procedure
  '
    call Sleep (1000)

end sub

function EnumWindowsProc(byVal hWnd as long, byVal lParam as long) as long ' {

    dim windowText  as string
    dim windowClass as string * 256
    dim retVal      as long
    
    windowText = space(GetWindowTextLength(hWnd) + 1)
    retVal     =       GetWindowText(hWnd, windowText, len(windowText))
    windowText = left$(windowText, retVal)

    if inStr(1, windowText, captionPart_, vbTextCompare) then

       call ShowWindow(hWnd, SW_SHOW)
       call SetForeGroundWindow(hWnd)

     '
     ' We have found a Window, the iteration
     ' process can be stopped
     '
       EnumWindowsProc = false
       exit function

    end if

    EnumWindowsProc = true

end function ' }
