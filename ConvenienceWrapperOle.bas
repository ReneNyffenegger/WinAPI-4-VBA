option explicit

function CoCreateGuid_ as GUID ' {

    if CoCreateGuid(CoCreateGuid_) <> 0 then
       MsgBox "Something went wrong with CoCreateGuid"
    end if

end function ' }

function StringFromGUID2_(rguid as GUID) as string ' {

    StringFromGUID2_ = space$(38)

    call StringFromGUID2 (rguid, strPtr(StringFromGUID2_), 38*2)

end function ' }
