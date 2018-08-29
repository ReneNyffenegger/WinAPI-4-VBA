option explicit


sub main() ' {

    const maxNumOfModules  = 1024
    const sizeOfHANDLE     =    4
    dim   bytesNeeded   as long

    dim bFilterModule   as Boolean
    dim lCountMatching  as long
    dim bRetVal         as Boolean
    dim nofModules      as long
    dim stringMaxPath   as string * MAX_PATH
    dim moduleBaseName  as string
    dim modulePath      as string
    dim hProc           as long
    dim module_info     as MODULEINFO
    dim lenString       as long

    dim moduleHandles(maxNumOfModules)     as long


    hProc = GetCurrentProcess()
    if hProc = 0 then
       debug.print "Failed to find current process"
       exit sub
    end if


    if EnumProcessModules(hProc, moduleHandles(0), (maxNumOfModules * sizeOfHANDLE), bytesNeeded) = false then
       debug.print "EnumProcessModules failed"
       exit sub
    end if


    nofModules = bytesNeeded / sizeOfHANDLE


    dim m as long
    for m = 0 to nofModules - 1

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

        debug.print modulePath
        debug.print "  " & moduleBaseName
        debug.print "  " & module_info.lpBaseOfDll
        debug.print "  " & module_info.SizeofImage
        debug.print "  " & module_info.EntryPoint

      skipModule:
    next

    erase moduleHandles
    call CloseHandle(hProc)

end sub ' }
