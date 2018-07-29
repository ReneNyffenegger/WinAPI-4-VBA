'
'  https://github.com/michaelneu/webxcel was very helpful
'
option explicit

public const AF_INET                  =      2
public const FD_SETSIZE               =     64
public const INADDR_ANY               =      0
public const SOCK_STREAM              =      1
public const SOL_SOCKET               =  65535
Public Const SO_RCVTIMEO              = &H1006
public const WSADESCRIPTION_LEN       =    256
public const WSASYS_STATUS_LEN        =    128
public const WSADESCRIPTION_LEN_ARRAY = WSADESCRIPTION_LEN + 1
public const WSASYS_STATUS_LEN_ARRAY  = WSASYS_STATUS_LEN  + 1

type WSADATA ' {
     wVersion       as integer
     wHighVersion   as integer
     szDescription  as string * WSADESCRIPTION_LEN_ARRAY
     szSystemStatus as string * WSASYS_STATUS_LEN_ARRAY
     iMaxSockets    as integer
     iMaxUdpDg      as integer
     lpVendorInfo   as string
end type ' }

type IN_ADDR ' {
     s_addr         as long
end type ' }

type fd_set ' {
     fd_count             as integer
     fd_array(FD_SETSIZE) as long
end type ' }

type timeval ' {
     tv_sec         as long
     tv_usec        as long
end type ' }

type sockaddr ' {
     sa_family      as integer
     sa_data        as string * 14
end type ' }

type sockaddr_in ' {
     sin_family     as integer
     sin_port       as integer
     sin_addr       as IN_ADDR
     sin_zero       as string * 8
end type ' }

' {
declare Function WSAStartup      lib "wsock32.dll"                (byVal versionRequired as long, wsa as WSADATA) as long
declare Function WSAGetLastError lib "wsock32.dll"                () as long
declare Function WSACleanup      lib "wsock32.dll"                () as long
declare Function socket          lib "wsock32.dll"                (byVal addressFamily as long, byVal socketType as long, byVal protocol as long) as long
declare Function htons           lib "wsock32.dll"                (byVal hostshort as long) as integer
declare Function bind            lib "wsock32.dll"                (byVal socket as long, name as sockaddr_in, byVal nameLength as integer) as long
declare Function listen          lib "wsock32.dll"                (byVal socket as long, byVal backlog as integer) as long
declare Function select_         lib "wsock32.dll" alias "select" (byVal nfds as integer, readfds as fd_set, writefds as fd_set, exceptfds as fd_set, timeout as timeval) as integer
declare Function accept          lib "wsock32.dll"                (byVal socket as long, clientAddress as sockaddr, clientAddressLength as integer) as long
declare Function setsockopt      lib "wsock32.dll"                (byVal socket as long, byVal level as long, byVal optname as long, ByRef optval as long, byVal optlen as integer) as long
declare Function send            lib "wsock32.dll"                (byVal socket as long, buffer as string, byVal bufferLength as long, byVal flags as long) as long
declare Function recv            lib "wsock32.dll"                (byVal socket as long, byVal buffer as string, byVal bufferLength as long, byVal flags as long) as long
declare Function closesocket     lib "wsock32.dll"                (byVal s as long) as long
' }

public sub FD_ZERO(byRef s as fd_set) ' {
    s.fd_count = 0
end sub ' }

sub FD_SET_(byVal fd as long, byRef s as fd_set) ' {
    dim i as integer
    i = 0

    do while i < s.fd_count
        if s.fd_array(i) = fd Then
            exit do
        end if

        i = i + 1
    loop

    if i = s.fd_count then
        if s.fd_count < FD_SETSIZE then
            s.fd_array(i) = fd
            s.fd_count = s.fd_count + 1
        end if
    end if

end sub ' }
