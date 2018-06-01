type POINTAPI
    x as long
    y as long
end type

type MSG
    hWnd    as Long
    message as Long
    wParam  as Long
    lParam  as Long
    time    as Long
    pt      as POINTAPI
end type

declare function WaitMessage lib "user32" () as long

declare function PeekMessage lib "user32" alias "PeekMessageA" ( _
     byRef lpMsg         as MSG , _
     byVal hwnd          as long, _
     byVal wMsgFilterMin as long, _
     byVal wMsgFilterMax as long, _
     byVal wRemoveMsg    as long) as long

declare function TranslateMessage lib "user32"(byRef lpMsg as MSG) as long

declare function PostMessage lib "user32" alias "PostMessageA" ( _
     byVal hwnd   as long, _
     byVal wMsg   as long, _
     byVal wParam as long, _
           lParam as any) as long

declare function FindWindow lib "user32" alias "FindWindowA" ( _
     byVal lpClassName  as string, _
     byVal lpWindowName as string) as long


public const  WM_CHAR    as long = &H102
public const  WM_KEYDOWN as long = &H100

public const  PM_REMOVE  as long = &H1
