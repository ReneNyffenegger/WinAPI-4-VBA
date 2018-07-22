option explicit

sub main()

    dim errNo  as long
    dim rc     as long
    dim langId as long

    dim errStr as string * FORMAT_MESSAGE_TEXT_LEN

  '
  ' The error number for which the text has to be retrieved:
  '
    errNo  = 1428

  '
  ' The language into which to format the message. 0 = default language.
  '
    langId = 0

    rc = FormatMessage (                                                _
           FORMAT_MESSAGE_FROM_SYSTEM or FORMAT_MESSAGE_IGNORE_INSERTS, _
           0                                                          , _
           errNo                                                      , _
           langId                                                     , _
           errStr                                                     , _
           FORMAT_MESSAGE_TEXT_LEN                                    , _
           0)

    if rc <> 0 then
       msgBox "Error text for " & errNo & " is: " & errStr
    else
       msgBox "Could not format message"
    end if

end sub
