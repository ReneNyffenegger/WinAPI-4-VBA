sub main()

   dim lLcid   as long
   dim LCData  as string
   dim locTxt  as string
   dim rc      as long

   lLcid = GetSystemDefaultLangID()

   LCData = string$(1024, 0)
 ' rc = GetLocaleInfo(lLcid, LOCALE_ILANGUAGE, LCData, len(LCData))
   rc = GetLocaleInfo(lLcid, LOCALE_SNAME    , LCData, len(LCData)) ' Get a string like "de-CH".

   if rc > 0 then
    ' Copy interesting part of LCData. Length is returned in rc, including
    ' \x0 character, therefore substracting 1.
      locTxt = left$(LCData, rc-1)
      debug.print "LOCALE_SNAME = " & locTxt
   else
      debug.print "Error occured: " & GetLastError()
   end if
end sub
