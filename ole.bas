option explicit


declare function CoTaskMemAlloc        lib "ole32"        ( _
        byVal cb        as long) as long

declare sub      CoTaskMemFree         lib "ole32"        ( _
        byVal pv        as long)

declare function IIDFromString         lib "ole32"        ( _
        byVal lpsz      as long, _
        byVal lpiid     as long) as long

declare function SysAllocStringByteLen lib "oleaut32"     ( _
        byVal psz       as long, _
        ByVal cblen     as long) as long

declare function VariantCopy           lib "oleaut32"     ( _
        byVal pvargDest as long, _
        byRef pvargSrc  as variant) as long
