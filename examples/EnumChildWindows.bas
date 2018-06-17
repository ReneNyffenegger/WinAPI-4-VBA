option explicit

global row_    as long
global indent_ as long
global cbFunc  as long
const  maxIter = 3333
global cntIter as long

'
'      The variable isTopLevel is set to true in the
'      call back that iterates over the top level
'      windows (CallBackFuncTopLevel).
'      In the function that iterates over a top
'      level's descendants, it is set to false.
'      Thus, it was needed to determine if in a recursive
'      strategy, I have to descend into the windows.
'      I am not sure, if the variable is still needed...
' 
global isTopLevel as boolean

sub main()
    row_    = 1
    indent_ = 1

    cbFunc = getAddr(addressOf CallBackFunc)

  '
  ' Iterate over top level windows (?)
  '
    call EnumWindows(addressOf CallBackFuncTopLevel, 0)

end sub

function getAddr(addr as long) as long
     getAddr = addr
end function

function CallBackFuncTopLevel(byVal hWnd as long, byVal lParam as long) as long

     isTopLevel = true
     call nextWindow(hWnd)

     CallBackFuncTopLevel = true

end function

function CallBackFunc(byVal hWnd as long, byVal lParam as long) as long


    dim hWndParent  as long

    if cntIter > maxIter then
       CallBackFunc = false
       exit function
    end if

  '
  ' Apparently, EnumChildWindows also iterates over GrandChildren
  ' and further descendants.
  ' We're interested in immediate children only:
  '
    hWndParent     = GetParent(hWnd)
    if not isTopLevel and hWndParent <> lParam then
       CallBackFunc = true
       exit function
    end if

    cntIter = cntIter + 1
  '
  ' We're pretty sure that we're not a top level window:
    isTopLevel = false

    call nextWindow(hWnd)


  '
  ' Return true to indicate that we want to continue
  ' with the enumeration of the windows:
  '
    CallBackFunc = true

end function

sub nextWindow(hWnd as long)

    dim hWndParent  as long
    dim windowText  as string
    dim windowClass as string * 256
    dim retVal      as long

    hWndParent     = GetParent(hWnd)

    cells(row_, 1) = hWndParent
    cells(row_, 2) = hWnd
'   cells(row_, 3) = lParam

    retVal = GetClassName(hWnd, windowClass, 255)
    windowClass = left$(windowClass, retVal)
    cells(row_, 4) = rtrim(windowClass)

    windowText = space(GetWindowTextLength(hWnd) + 1)
    retVal     =       GetWindowText(hWnd, windowText, len(windowText))
    windowText = left$(windowText, retVal)
    cells(row_, 5 + indent_) = ">" & windowText

    row_ = row_ + 1
    indent_ = indent_ + 1
    call EnumChildWindows(hWnd, cbFunc, hWnd)
    indent_ = indent_ - 1

end sub
