option explicit

sub main() ' {

    dim hProc as long

    hProc = GetCurrentProcess()
    if hProc = 0 then
       debug.print "Failed to find current process"
       exit sub
    end if

    call enumProcModules(hProc, environ$("TEMP") & "\enumProcModules.txt")

end sub ' }

sub enumProcModules(hProc as long, optional fileName as string) ' {

    const maxNumOfModules  = 1024
    const sizeOfHANDLE     =    4
    dim   bytesNeeded   as long

    dim bFilterModule   as boolean
    dim lCountMatching  as long
    dim bRetVal         as boolean
    dim nofModules      as long
    dim stringMaxPath   as string * MAX_PATH
    dim moduleBaseName  as string
    dim modulePath      as string
    dim module_info     as MODULEINFO
    dim lenString       as long
    dim fileNo          as integer

    dim moduleHandles(maxNumOfModules)     as long

    if EnumProcessModules(hProc, moduleHandles(0), (maxNumOfModules * sizeOfHANDLE), bytesNeeded) = false then
       debug.print "EnumProcessModules failed"
       exit sub
    end if

    nofModules = bytesNeeded / sizeOfHANDLE
    
    if fileName <> "" then
       fileNo = freeFile()
       open fileName for output as #fileNo
    end if

    dim m as long
    for m = 0 to nofModules - 1 ' {

        if moduleHandles(m) = 0 then
           goto skipModule
        end if

        if GetModuleInformation(hProc, moduleHandles(m), module_info, bytesNeeded) = 0 then
           debug.print "Could not get module info"
           goto skipModule
        end if

        lenString = GetModuleFileNameEx(hProc, moduleHandles(m), stringMaxPath, MAX_PATH)
        modulePath = mid$(stringMaxPath, 1, lenString)

        lenString = GetModuleBaseName(hProc, moduleHandles(m), stringMaxPath, MAX_PATH)
        moduleBaseName = mid$(stringMaxPath, 1, lenString)

            if fileName = "" then
               debug.print modulePath
               debug.print "  " & moduleBaseName
               debug.print "  " & module_info.lpBaseOfDll
               debug.print "  " & module_info.SizeofImage
               debug.print "  " & module_info.EntryPoint
            else
               print# fileNo, modulePath
               print# fileNo, "  " & moduleBaseName
               print# fileNo, "  " & module_info.lpBaseOfDll
               print# fileNo, "  " & module_info.SizeofImage
               print# fileNo, "  " & module_info.EntryPoint
            end if
    
          skipModule:
    next m ' }

    if fileName <> "" then
       close# fileNo
    end if

    erase moduleHandles
    call CloseHandle(hProc)

end sub ' }
