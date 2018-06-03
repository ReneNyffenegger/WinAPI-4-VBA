option explicit

sub main()

   dim userName as string * 32
   dim result   as long

   result = GetUserName(userName, 32)

   if result <> 0 then
      msgBox "The currently logged in user is " & userName
   else
      msgBox "Calling GetUserName failed"
   end if

end sub
