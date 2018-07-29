option explicit

dim keyboardLayout as long

sub main()


    keyboardLayout = GetKeyboardLayout(0)

    showCharInfo "c"
    showCharInfo "C"
    showCharInfo "="
    showCharInfo "+"
'   showCharInfo chr(1)


end sub

sub showCharInfo(char as string) ' {

    dim keyScan as integer
    dim shift   as boolean
    dim ctrl    as boolean
    dim vkKey   as integer

    keyScan = VkKeyScanEx(asc(char), keyboardLayout)

    shift = keyScan and &h100
    ctrl  = keyScan and &h200
    vkKey = keyScan and &h0ff

    debug.print char & ": " + iif(shift, "shift ", "      ") + iif(ctrl, "ctrl ", "     ") & chr(vkKey) & " " & vkKey


end sub ' }
