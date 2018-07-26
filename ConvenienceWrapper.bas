option explicit

global g_hWnd       as long
global captionPart_ as string
global className_   as string
global windowText_  as string

sub clipboardPutText(txt as string) ' {
   dim memory          as long
   dim lockedMemory    as long

   memory       = GlobalAlloc(GHND, len(txt) + 1)
   if memory = 0 then
      msgBox "GlobalAlloc failed"
      exit sub
   end if

   lockedMemory = GlobalLock(memory)
   if lockedMemory = 0 then
      msgBox "GlobalLock failed"
      exit sub
   end if

   lockedMemory = lstrcpy(lockedMemory, txt)

   call GlobalUnlock(memory)

   if openClipboard(0) = 0 Then
      msgBox "openClipboard failed"
      exit sub
   end if

   call EmptyClipboard()

   call SetClipboardData(CF_TEXT, memory)

   if CloseClipboard() = 0 then
      msgBox "CloseClipboard failed"
   end if
end sub ' }

function GetWindowText_(hWnd as long) as string ' {
    dim retVal      as long

    GetwindowText_ = space(GetWindowTextLength(hWnd) + 1)
    retVal         = GetWindowText(hWnd, GetWindowText_, len(GetwindowText_))
    GetWindowText_ = left$(GetwindowText_, retVal)

end function ' }

function GetClassName_(hWnd as long) as string ' {
    dim windowClass as string * 256
    dim retVal      as long

    retVal = GetClassName(hWnd, windowClass, 255)
    GetClassName_ = left(windowClass, retVal)

end function ' }

function GetComputerName_() as string ' {
   dim NetBiosName as string * 32
   dim rc      as long
   dim length  as long

   length = 32
   rc = GetComputerName(NetBiosName, length)

   if rc <> 0 then
      GetComputerName_ = left$(NetBiosName, length)
      exit function
   end if

   GetComputerName_ = "?"

end function ' }

function GetLastErrorText() as string ' {
    dim errNo  as long
    dim rc     as long
    dim langId as long

    errNo = GetLastError()
    GetLastErrorText = string$(FORMAT_MESSAGE_TEXT_LEN, chr(0))

    langId = 0

    rc = FormatMessage (                                                _
           FORMAT_MESSAGE_FROM_SYSTEM or FORMAT_MESSAGE_IGNORE_INSERTS, _
           0                                                          , _
           errNo                                                      , _
           langId                                                     , _
           GetLastErrorText                                           , _
           FORMAT_MESSAGE_TEXT_LEN                                    , _
           0)

    if rc = 0 then
       GetLastErrorText = errNo
    end if
end function ' }

function FindWindow_WindowNameContains(captionPart as string) as long ' {
    captionPart_ = captionPart
    call EnumWindows(addressOf FindWindow_WindowNameContains_cb, byVal 0&)
    FindWindow_WindowNameContains = g_hWnd
end function ' }

function FindWindow_WindowNameContains_cb(byVal hWnd as long, byVal lParam as long) as long ' {

    dim windowText  as string
    dim windowClass as string * 256
    dim retVal      as long

    windowText = space(GetWindowTextLength(hWnd) + 1)
    retVal     =       GetWindowText(hWnd, windowText, len(windowText))
    windowText = left$(windowText, retVal)

    if inStr(1, windowText, captionPart_, vbTextCompare) then

       g_hWnd = hWnd

     '
     ' We have found a Window, the iteration
     ' process can be stopped
     '
       FindWindow_WindowNameContains_cb = false
       exit function

    end if

    FindWindow_WindowNameContains_cb = true

end function ' }

function FindWindow_ClassName(className as string) as long ' {
    captionPart_ = className
    g_hWnd = 0
    call EnumWindows(addressOf FindWindow_ClassName_cb, byVal 0&)
    FindWindow_ClassName = g_hWnd
end function ' }

function FindWindow_ClassName_cb(byVal hWnd as long, byVal lParam as long) as long ' {

    dim windowText  as string
    dim windowClass as string * 256
    dim retVal      as long

    if GetClassName_(hWnd) = captionPart_ then

       g_hWnd = hWnd
       FindWindow_ClassName_cb = false
       exit function

    end if

    FindWindow_ClassName_cb = true

end function ' }

function FindWindow_ClassName_WindowText(className as string, windowText as string) as long ' {
    className_  = className
    windowText_ = windowText
    g_hWnd = 0
    call EnumWindows(addressOf FindWindow_ClassName_WindowText_cb, byVal 0&)
    FindWindow_ClassName_WindowText = g_hWnd
end function ' }

function FindWindow_ClassName_WindowText_cb(byVal hWnd as long, byVal lParam as long) as long ' {

    if GetClassName_(hWnd) = className_ and GetWindowText_(hWnd) = windowText_  then

       g_hWnd = hWnd
       FindWindow_ClassName_WindowText_cb = false
       exit function

    end if

    FindWindow_ClassName_WindowText_cb = true

end function ' }

function GetWindowRect_(hWnd as long) as RECT ' {
    dim r as RECT
    GetWindowRect hWnd, r
    GetWindowRect_ = r
end function ' }

sub shellOpen(path as string) ' {
    ShellExecute 0, "Open", path, "", "", 1
end sub ' }

sub SendInputText(text as string) ' {
    dim i         as long
    dim c         as string
'   dim input_    as INPUT_
    dim input_(1) as INPUT_
    dim sizeINPUT as long

    sizeINPUT = lenB(input_(0))

    input_(0).dwType = INPUT_KEYBOARD

    for i = 0 to len(text) - 1 ' {
        c = mid(text, i+1, 1)

        debug.print "c = " & c

        input_(0).dwFlags = 0

        if c >= "A" and c<= "Z" then ' {
           input_(0).wVK = VK_LSHIFT
           SendInput 1, input_(0), sizeINPUT
        end if ' }


        input_(0).wVK = VkKeyScan(asc(lcase(c)))
        SendInput 1, input_(0), sizeINPUT

        input_(0).dwFlags = KEYEVENTF_KEYUP
        SendInput 1, input_(0), sizeINPUT

        if c >= "A" and c<= "Z" then ' {
           input_(0).wVK = VK_LSHIFT
           SendInput 1, input_(0), sizeINPUT
        end if ' }

    next i ' }
end sub ' }
