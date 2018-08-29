option explicit

sub main()
    EnumDisplayMonitors 0, 0, addressOf monitorEnumProc, 0
end sub


function monitorEnumProc(byVal hMonitor as long, byVal hdcMonitor as long, byRef rMonitor as RECT, byVal dwData as long) as long
   debug.print "Found monitor with top left = " & rMonitor.left & "," & rMonitor.top & " and size = " & (rMonitor.right-rMonitor.left) & "x" (rMonitor.bottom-rMonitor.top)

 ' Return true so as to continue with the enumeration
   monitorEnumProc = true
end function
