'
'  https://stackoverflow.com/questions/11153995/is-there-any-event-that-fires-when-keys-are-pressed-when-editing-a-cell
'
'  call trackKeyPressInit
'  call stopKeyWatch
'
'  TODO: Win 64 Api -> https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/ptrsafe-keyword
'
option explicit

private bExitLoop as boolean

sub trackKeyPressInit()

    dim msgMessage           as MSG
    dim propagateKeyPress    as boolean
    dim iKeyCode             as integer
    dim lXLhwnd              as long

    on error GoTo errHandler:
    
    application.enableCancelKey = xlErrorHandler
        
  ' initialize this boolean flag.
    bExitLoop = False
        
  ' get the app hwnd.
    lXLhwnd = FindWindow("XLMAIN", application.caption)
        
    do
        WaitMessage
        
      ' Check for a key press and remove it from the msg queue.
      
        if PeekMessage (msgMessage, lXLhwnd, WM_KEYDOWN, WM_KEYDOWN, PM_REMOVE) Then
        
          ' strore the virtual key code for later use.
            iKeyCode = msgMessage.wParam
            
          ' translate the virtual key code into a char msg.
            TranslateMessage msgMessage
            call PeekMessage(msgMessage, lXLhwnd, WM_CHAR, WM_CHAR, PM_REMOVE)
            
          ' for some obscure reason, the following
          ' keys are not trapped inside the event handler
          ' so we handle them here.
          
            if iKeyCode = vbKeyBack   then SendKeys "{BS}"
            if iKeyCode = vbKeyReturn then SendKeys "{ENTER}"
            
            propagateKeyPress = true
            
          ' The VBA RaiseEvent statement does not seem to return ByRef arguments
          ' so we call a KeyPress routine rather than a propper event handler.
            call keyPressed(byVal msgMessage.wParam, byVal iKeyCode, propagateKeyPress)
            
          ' if the key pressed is allowed post it to the application.
            if propagateKeyPress then
               call PostMessage(lXLhwnd, msgMessage.Message, msgMessage.wParam, 0)
            end if
        end if
    errHandler:
      ' allow the processing of other msgs.
        doEvents
    loop until bExitLoop

end sub

sub stopKeyWatch()
    bExitLoop = true
end sub



private sub keyPressed(byVal KeyAscii          as integer, _
                       byVal KeyCode           as integer, _
                       byRef propagateKeyPress as boolean)

    debug.print keyAscii
    
    if chr(keyAscii) = "9" then
       bExitLoop = true
    end if
    
    if chr(keyAscii) = "8" then
       propagateKeyPress = false
    end if
       

 '  if Not Intersect(Target, Range("A1:D10")) Is Nothing Then
 '      if Chr(KeyAscii) Like "[0-9]" Then
 '          MsgBox MSG & Range("A1:D10").Address(False, False) _
 '          & """ .", vbCritical, TITLE
 '          Cancel = True
 '      end If
 '  end If

End Sub

