option explicit

sub main() ' {

    dim secAttrPrc as SECURITY_ATTRIBUTES : secAttrPrc.nLength = len(secAttrPrc)
    dim secAttrThr as SECURITY_ATTRIBUTES : secAttrThr.nLength = len(secAttrThr)

    dim startInfo  as STARTUPINFO
    dim procInfo   as PROCESS_INFORMATION

    if CreateProcess (                                         _
         lpApplicationName      :=   vbNullString            , _
         lpCommandLine          :=  "cmd.exe"                , _
         lpProcessAttributes    :=   secAttrPrc              , _
         lpThreadAttributes     :=   secAttrThr              , _
         bInheritHandles        :=   false                   , _
         dwCreationFlags        :=   0                       , _
         lpEnvironment          :=   0                       , _
         lpCurrentDriectory     :=   environ("USERPROFILE")  , _
         lpStartupInfo          :=   startInfo               , _
         lpProcessInformation   :=   procInfo )  then

     else
        MsgBox "Couldn't create process"
     end if

end sub ' }
