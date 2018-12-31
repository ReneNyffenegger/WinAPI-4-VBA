option explicit

enum HRESULT_values
  ' NULL_          =          0
    S_OK           =          0 ' Operation successful
    S_FALSE        =          1
    E_ABORT        = &h80004004 ' Operation aborted
    E_ACCESSDENIED = &h80070005 ' General access denied error
    E_FAIL         = &h80004005 ' Unspecified failure
    E_HANDLE       = &h80070006 ' Invalid handle
    E_INVALIDARG   = &h80070057 ' One or more arguments are not valid
    E_NOINTERFACE  = &h80004002 ' No such interface supported
    E_NOTIMPL      = &h80004001 ' Not implemented
    E_OUTOFMEMORY  = &h8007000E ' Failed to allocate necessary memory
    E_POINTER      = &h80004003 ' Invalid pointer
    E_UNEXPECTED   = &h8000FFFF ' Unexpected failure
end enum

type GUID ' {
  '
  '  Declared in rpcdce.h / included by rpc.h
  '
     Data1          as long
     Data2          as integer
     Data3          as integer
     Data4 (0 to 7) as byte
end  type ' }

declare function CoTaskMemAlloc        lib "ole32"        ( _
        byVal cb        as long) as long

declare function CoCreateGuid          lib "ole32"        ( _
              pguid     as GUID) as long

declare sub      CoTaskMemFree         lib "ole32"        ( _
        byVal pv        as long)

declare function IIDFromString         lib "ole32"        ( _
        byVal lpsz      as long, _
        byVal lpiid     as long) as long

declare function StringFromGUID2       lib "ole32"        ( _
              rguid     as GUID, _
        byVal lpOleChar as any , _
        byVal cbmax     as long) as long

declare function SysAllocStringByteLen lib "oleaut32"     ( _
        byVal psz       as long, _
        ByVal cblen     as long) as long

declare function VariantCopy           lib "oleaut32"     ( _
function CoCreateGuid_ as GUID ' {

    if CoCreateGuid(CoCreateGuid_) <> 0 then
       MsgBox "Something went wrong with CoCreateGuid"
    end if

end function ' }

function StringFromGUID2_(rguid as GUID) as string ' {

    StringFromGUID2_ = space$(38)

    call StringFromGUID2 (rguid, strPtr(StringFromGUID2_), 38*2)

end function ' }
        byVal pvargDest as long, _
        byRef pvargSrc  as variant) as long
