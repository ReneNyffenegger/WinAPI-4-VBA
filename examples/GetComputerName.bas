option explicit

sub main()

   dim NetBiosName as string * 32
   dim result      as long

   result = GetComputerName(NetBiosName, 32)

   if result <> 0 then
      msgBox "The Net BIOS name of the computer is " & NetBiosName
   else
      msgBox "Calling GetComputerName failed"
   end if

end sub
