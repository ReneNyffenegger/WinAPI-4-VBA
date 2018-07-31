option explicit

sub main() ' {

    debug.print "Number of display monitors on the desktop: " & GetSystemMetrics(SM_CMONITORS  )
    debug.print "Width of primary monitor:                  " & GetSystemMetrics(SM_CXSCREEN   )
    debug.print "Height of primary monitor:                 " & GetSystemMetrics(SM_CYSCREEN   )
    debug.print "Slow (low-end) processor:                  " & GetSystemMetrics(SM_SLOWMACHINE)

end sub ' }
