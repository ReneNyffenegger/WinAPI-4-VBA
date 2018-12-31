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
        byVal pvargDest as long, _
        byRef pvargSrc  as variant) as long
