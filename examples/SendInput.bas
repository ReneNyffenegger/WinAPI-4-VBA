option explicit

dim keyInput()    as INPUT_
dim sizeINPUT     as long

sub main() ' {

    sizeINPUT = lenB(keyInput(0))

    call Windows_R()
    call Sleep(100)

    call pressAndRelease("notepad")
    call enter()

    call Sleep(200)

    pressAndRelease ("someText")

end sub ' }

sub Windows_R() ' {

    redim keyInput(0 to 3)

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

end sub ' }

sub pressAndRelease(text as string) ' {

    dim c as string
    dim i as long
    redim keyInput(0 to len(text)*2 - 1)    as INPUT_

    for i = 0 to len(text) -1 ' {

        c = mid(text, i+1, 1)

        keyInput(i*2).dwType   = INPUT_KEYBOARD
        keyInput(i*2).dwFlags  = 0
        keyInput(i*2).wVK      = VkKeyScan(asc(c))

        keyInput(i*2+1).dwType   = INPUT_KEYBOARD
        keyInput(i*2+1).dwFlags  = KEYEVENTF_KEYUP
        keyInput(i*2+1).wVK      = VkKeyScan(asc(c))

    next i ' }

    call SendInput(i*2, keyInput(0), sizeINPUT)

end sub ' }

sub enter() ' {
    redim keyInput(0 to 1) '

    keyInput(0).dwType  = INPUT_KEYBOARD
    keyInput(0).dwFlags = 0        ' Press key
    keyInput(0).wVK     = VK_RETURN

    keyInput(1).dwType  = INPUT_KEYBOARD
    keyInput(1).dwFlags = KEYEVENTF_KEYUP
    keyInput(1).wVK     = VK_RETURN

    call SendInput(2, keyInput(0), sizeINPUT)

end sub ' }
