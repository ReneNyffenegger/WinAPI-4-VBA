option explicit

sub main()

    dim hWndNotepad     as long

  '
  ' Start notepad
  '
    ShellExecute 0, "Open", "notepad.exe", "", "", 1

  '
  ' Wait for notepad to initialize.
  '
    Sleep 200

  '
  ' Find the window handle for notepad
  '
    hWndNotepad = FindWindow("notepad", vbNullString)

    dim r as RECT
    call GetWindowRect(hWndNotepad, r)
    msgBox "Notepad dimensions: " & (r.right - r.left) & " x " & (r.bottom - r.top) & ", top left corner at " & r.left & "/" & r.top


end sub
