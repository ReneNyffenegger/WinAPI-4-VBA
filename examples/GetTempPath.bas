option explicit

sub main()

   dim tempPath as string * 512
   dim result   as long

   result = GetTempPath(512, tempPath)

   if result <> 0 then
      msgBox "Temp path: " & tempPath
   else
      msgBox "Calling GetTempPath failed"
   end if

end sub

