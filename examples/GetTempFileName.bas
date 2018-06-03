sub main()

   dim tempPath as string * 512
   dim tempFile as string * 512 ' Should be MAX_PATH !
   dim uUnique  as long
   

   call GetTempPath(512, tempPath)

   uUnique = 0 ' Create unique temporary file names
   call GetTempFileName(tempPath, "abc", uUnique, tempFile)

   msgBox "The temporary file is " & tempFile

   open tempFile for output as 1

   print# 1, "Foo bar baz"
   print# 1, "One two three"

   close# 1

end sub
