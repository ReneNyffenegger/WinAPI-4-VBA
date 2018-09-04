'
' http://www.tech-archive.net/Archive/VB/microsoft.public.vb.general.discussion/2007-09/msg00228.html
'
option explicit

'
' Hmmm… why is this necessary?
'
private declare sub MoveMemoryLong lib "kernel32" alias "RtlMoveMemory" ( _
        Target   as Any , _
  byVal LPointer as long, _
  byVal cbCopy   as long)


sub main() 
  dim IM       as LOADED_IMAGE
  dim peHeader as IMAGE_PE_FILE_HEADER
  dim expTable as IMAGE_EXPORT_DIRECTORY_TABLE

  dim dllPath as string: dllPath = "C:\Windows\System32\msvbvm60.dll"

  if MapAndLoad(dllPath, dllPath, IM, true, true) = 0 then
     debug.print "Could nod map " & dllPath 
     exit sub
  end if

  on error goto err_

' Copy pe file header:
  RtlMoveMemory byVal varPtr(peHeader), byVal IM.pFileHeader, 256

' Get export table offset as relative virtual address (RVAs)
  dim expRVA as long
  expRVA = peHeader.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).RVA

  if expRVA = 0 then
     call UnMapAndLoad(IM)
     
  end if

' Convert RVA to VA:
  dim expVA as long
  expVA = ImageRvaToVa(IM.pFileHeader, IM.MappedAddress, expRVA, 0&)

  RtlMoveMemory byVal varPtr(expTable), byVal expVA, LenB(expTable)

  dim nofExports as long
  nofExports = expTable.NumberOfNames

  dim ptrExportedFuncs as long ' , NamePt as long
  ptrExportedFuncs = ImageRvaToVa(IM.pFileHeader, IM.MappedAddress, expTable.pAddressOfNames, 0&)

  dim ptrName as long
  ptrName = ptrExportedFuncs ' first pointer in array

  dim ptrNameCopy as long
  MoveMemoryLong ptrNameCopy, ptrName, 4

  dim i as long
  for i = 0 to nofExports - 1
      ptrNameCopy = ImageRvaToVa(IM.pFileHeader, IM.MappedAddress, ptrNameCopy, 0&)
      debug.print LPSTRtoBSTR(ptrNameCopy)
      ptrName = (ptrName + 4)
      MoveMemoryLong ptrNameCopy, ptrName, 4
  next i

  call UnMapAndLoad(IM)
  exit sub

err_:
  call UnMapAndLoad(IM)
  debug.print "Error occured"
end sub

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
