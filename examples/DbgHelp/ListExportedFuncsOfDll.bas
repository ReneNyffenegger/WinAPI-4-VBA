'
' http://www.tech-archive.net/Archive/VB/microsoft.public.vb.general.discussion/2007-09/msg00228.html
'
option explicit

'
' Hmmm… why is this necessary?
'
private declare sub MoveMemoryLong lib "kernel32" alias "RtlMoveMemory" ( _
        Target   as any , _
  byVal LPointer as long, _
  byVal cbCopy   as long)


sub listExportedFuncs() ' {
  dim img      as LOADED_IMAGE
  dim peHeader as IMAGE_NT_HEADERS32 '  Was IMAGE_PE_FILE_HEADER
  dim expTable as IMAGE_EXPORT_DIRECTORY_TABLE

  dim dllPath as string
      dllPath = "C:\Windows\System32\msvbvm60.dll"
'     dllPath = "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7.1\VBE7.DLL"

  if MapAndLoad(dllPath, dllPath, img, true, true) = 0 then
     debug.print "Could not map " & dllPath 
     exit sub
  end if

  on error goto err_

  dim fOut as integer: fOut = freeFile
  open environ$("TEMP") & "\exportedFuncs.txt" for output as #fOut

' Copy PE file header:
  RtlMoveMemory byVal varPtr(peHeader), byVal img.FileHeader, lenb(peHeader)

' Get export table offset as relative virtual address (RVAs)
  dim expRVA as long
  expRVA = peHeader.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).RVA

  if expRVA = 0 then
     call UnMapAndLoad(img)
  end if

' Convert RVA to VA:
  dim expVA as long
  expVA = ImageRvaToVa(img.FileHeader, img.MappedAddress, expRVA, 0&)

  RtlMoveMemory byVal varPtr(expTable), byVal expVA, lenB(expTable)

  dim nofExports as long
  nofExports = expTable.NumberOfNames

  dim ptrToArrayOfExportedFuncNames     as long
  dim ptrToArrayOfExportedFuncAddresses as long

  ptrToArrayOfExportedFuncNames     = ImageRvaToVa(img.FileHeader, img.MappedAddress, expTable.AddressOfNames    , 0&)
  ptrToArrayOfExportedFuncAddresses = ImageRvaToVa(img.FileHeader, img.MappedAddress, expTable.AddressOfFunctions, 0&)

  dim i as long
  for i = 0 to nofExports - 1

      dim RVAfuncName    as long
      dim RVAfuncAddress as long

      MoveMemoryLong RVAfuncName   , ptrToArrayOfExportedFuncNames    , 4
      MoveMemoryLong RVAfuncAddress, ptrToArrayOfExportedFuncAddresses, 4

      dim VAfuncName    as long
      dim VAfuncAddress as long

      VAfuncName     = ImageRvaToVa(img.FileHeader, img.MappedAddress, RVAfuncName   , 0&)
      VAfuncAddress  = ImageRvaToVa(img.FileHeader, img.MappedAddress, RVAfuncAddress, 0&)

      print# fOut, VAfuncAddress & ": " & LPSTRtoBSTR(VAfuncName)
 
      ptrToArrayOfExportedFuncNames     = ptrToArrayOfExportedFuncNames     + 4
      ptrToArrayOfExportedFuncAddresses = ptrToArrayOfExportedFuncAddresses + 4
  next i

  call UnMapAndLoad(img)

  close# fOut
  exit sub

err_:
  call UnMapAndLoad(img)
  debug.print "Error occured: " & err.description
end sub ' listExportedFuncs }

private function LPSTRtoBSTR(byVal lpString as long) as string ' {
   dim lenS as long
   dim ptrToZero  as long

   lenS = lstrlen(lpString)

   LPSTRtoBSTR = string$(lenS, 0)
   RtlMoveMemory byVal StrPtr(LPSTRtoBSTR), byVal lpString, lenS

   LPSTRtoBSTR = StrConv(LPSTRtoBSTR, vbUnicode)

   ptrToZero = inStr(1, LPSTRtoBSTR, chr(0), 0)

   if ptrToZero > 0 then
      LPSTRtoBSTR = Left$(LPSTRtoBSTR, ptrToZero - 1)
   end if

end function ' }
