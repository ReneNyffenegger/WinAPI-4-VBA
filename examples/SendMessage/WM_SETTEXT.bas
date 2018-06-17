option explicit

dim hWndNotepadEdit as long

sub main()

    dim hWndNotepad     as long
    dim hWndNotepadEdit as long

  '
  ' Start notepad
  '
    ShellExecute 0, "Open", "notepad.exe", "", "", 1
  '
  ' Find the window handle for notepad
  '
    hWndNotepad = FindWindow("notepad", vbNullString)
    while hWndNotepad = 0
    '
    ' Wait a while if notepad window not yet initialized:
    '
      Sleep 100 
      hWndNotepad = FindWindow("notepad", vbNullString)
    wend

  '
  ' Find the window handle of the edit control
  ' in notepad:
  '
    hWndNotepadEdit = FindWindowEx(hWndNotepad, 0, "Edit", vbNullString)

  '
  ' Send WM_SETTEXT to the edit control.
  '
  ' Note,using PostMessage instead of SendMessage does not work.
  ' I have no clue why SendMessage works, but PostMessage does not!
  '
    SendMessage hWndNotepadEdit, WM_SETTEXT, 0, byVal "Some text that was sent with WM_SETTEXT"

end sub
