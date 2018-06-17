option explicit

dim hWndNotepadEdit as long

sub main()

    dim hWndNotepad     as long
    dim hWndNotepadEdit as long

  '
  ' Start notepad
  '
    ShellExecute 0, "Open", "notepad.exe", "", "", 1

    Sleep 200

  '
  ' Find the window handle for notepad
  '
    hWndNotepad = FindWindow("notepad", vbNullString)

  '
  ' Find the window handle of the edit control
  ' in notepad:
  '
    hWndNotepadEdit = FindWindowEx(hWndNotepad, 0, "Edit", vbNullString)

    debug.print "hWndNotepad     = " & hWndNotepad
    debug.print "hWndNotepadEdit = " & hWndNotepadEdit
    debug.print "Parent of edit  = " & GetParent(hWndNotepadEdit)

  '
  ' Simulate pressing VK_F5 to insert the current date:
  '
    PostMessage hWndNotepadEdit, WM_KEYDOWN, VK_F5, 1
    Sleep 50
    PostMessage hWndNotepadEdit, WM_KEYUP  , VK_F5, 1

    Sleep 2000

end sub
