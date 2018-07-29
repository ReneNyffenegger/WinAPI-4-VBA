'
'  https://github.com/michaelneu/webxcel was very helpful
'
option explicit

sub main() ' {

    if not initWinsock then
       msgBox "Could not initialize Winsock"
    end if

    debug.print "winsock initialized"

    dim serverSocket as long
    serverSocket = createServerSocket(8888)

    acceptConnections serverSocket

    closeSocket serverSocket
    WSACleanup

end sub ' }

function createServerSocket(byVal port as long) ' {

    createServerSocket = socket(AF_INET, SOCK_STREAM, 0)

    dim endPoint as sockaddr_in
    endPoint.sin_family       = AF_INET
    endPoint.sin_addr.s_addr  = INADDR_ANY
    endpoint.sin_port         = htons(port)

'   debug.print "lenB: " & lenB(endPoint)

    dim rc as long
    rc = bind(createServerSocket, endpoint, 16)
    if rc <> 0 then
       msgBox "Could not bind, error = " & WSAGetLastError()
       exit function
    end if

    rc = listen(createServerSocket, 10) ' 10 = backlog
    if rc <> 0 then
       msgBox "Could not listen"
    end if

end function ' }

sub acceptConnections(serverSocket as long) ' {

    dim clientSocket as long

    dim i as long
    i = 0

    do while i < 200
       i = i + 1
       sleep 100
       debug.print "i = " & i

       clientSocket = getClientSocket(serverSocket)

       if clientSocket = 0 then
          goto SKIP_THIS_ITERATION
       end if

       dim reqText as string
       reqText = getStringFromSocket(clientSocket)

       dim textResponse as string
       textResponse = "HTTP/1.1 200 OK" & chr(10)
       textResponse = textResponse & "Content-Type: text/html" & chr(10)
       textResponse = textResponse & chr(10)
       textResponse = textResponse & "<!doctype html>" & chr(10)
       textResponse = textResponse & "<html><body>Request was:<br><code><pre>"
       textResponse = textResponse & reqText
       textResponse = textResponse & "</pre></code></body></html>"

       send clientSocket, byVal textResponse, len(textResponse), 0

       closeSocket clientSocket

    SKIP_THIS_ITERATION:
    loop

end sub ' }

function getClientSocket(serverSocket as long) as long ' {
    dim      fdSet as fd_set
    dim emptyFdSet as fd_set
    dim rc         as integer

    FD_ZERO                  fdSet
    FD_SET_    serverSocket, fdSet

    dim timeOutMs as long
    timeOutMs = 500

    dim timeOut  as timeval
    timeOut.tv_sec  = timeOutMs  /  1000
    timeOut.tv_usec = timeOutMs Mod 1000

    rc = select_(serverSocket, fdSet, emptyFdSet, emptyFdSet, timeOut)
    if rc = 0 then
       getClientSocket = 0
       exit function
    end if

    dim socketAddress as sockaddr
    getClientSocket = accept(serverSocket, socketAddress, 16)

    if getClientSocket = -1 then
       getClientSocket = 0
       exit function
    end if

    rc = setsockopt(getClientSocket, SOL_SOCKET, SO_RCVTIMEO, timeOutMs, 4)

end function ' }

function getStringFromSocket(s as long) ' {
    dim message   as string
    dim buffer    as string * 1024
    dim readBytes as long

    message = ""

    do
        buffer = ""
        readBytes = recv(s, buffer, len(buffer), 0)

        if readBytes > 0 Then
           message = message & Trim(buffer)
        end if
    loop while readBytes > 0

    getStringFromSocket = trim(message)

end function ' }

function initWinsock() as boolean ' {

    dim wsaVersion as long
        wsaVersion = 257


    dim rc  as long
    dim wsa as WSADATA

    rc = WSAStartup(wsaVersion, wsa)

    if rc <> 0 then
       initWinsock = false
       exit function
    end if

    initWinsock = true

end function ' }
