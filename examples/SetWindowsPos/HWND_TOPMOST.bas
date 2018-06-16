option explicit

sub main()

    dim hWndNotepad as long

    call ShellExecute(0, "Open", "notepad.exe", ""    , "", 1)

  '
  ' Sleep a short while to give notepad.exe time to start and initialize.
  ' Otherwise, the following FindWindow would be too early and not find
  ' any window.
  '
    call Sleep(100)

  '
  ' The first argument of FindWindow expects a class name.
  ' For notepad, the class name is notepad.
  '
    hWndNotepad = FindWindow("notepad", vbNullString)

    if hWndNotepad = 0 then
       msgBox "could not find notepad window"
       exit sub
    end if

  '
  ' Put the 
    call SetWindowPos(hWndNotepad, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE)

end sub
