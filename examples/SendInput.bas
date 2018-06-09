option explicit


sub main()

    dim keyInput() as INPUT_

    dim sizeINPUT as long

    redim keyInput(0 to 3)

    sizeINPUT = lenB(keyInput(0))

    keyInput(0).dwType  = INPUT_KEYBOARD
    keyInput(0).dwFlags = 0        ' Press key
    keyInput(0).wVK     = VK_LWIN

    keyInput(1).dwType  = INPUT_KEYBOARD
    keyInput(1).dwFlags = 0        ' Press key
    keyInput(1).wVK     = VkKeyScan(asc("r"))

    keyInput(2).dwType  = INPUT_KEYBOARD
    keyInput(2).dwFlags = KEYEVENTF_KEYUP
    keyInput(2).wVK     = VkKeyScan(asc("r"))

    keyInput(3).dwType  = INPUT_KEYBOARD
    keyInput(3).dwFlags = KEYEVENTF_KEYUP
    keyInput(3).wVK     = VK_LWIN

    call SendInput(4, keyInput(0), sizeINPUT)


end sub


