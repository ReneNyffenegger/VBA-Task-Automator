option explicit

type keyEv ' {
     pressed    as boolean
     vkCode     as byte     ' VK_*
end type ' }

dim hookHandle  as long 
dim hookStarted as boolean

' dim ctrlIsDown as boolean
' dim altIsDown  as bool


const nofKeyEventsStored = 20
dim   lastKeyEvents(nofKeyEventsStored) as keyEv
dim   curKeyEvent as byte
dim   nextKeyEv  as keyEv

sub StartTaskAutomator() ' {

    call initLastKeyevents()
    nextKeyEv.pressed = false


    debug.print "TaskAutomator started"

    if hookStarted = false then
        hookHandle = SetWindowsHookEx(  _ 
           WH_KEYBOARD_LL                  , _
           addressOf LowLevelKeyboardProc  , _
           application.hInstance           , _
           0 )

        if hookHandle <> 0 then
           hookStarted = true
        else
           msgBox "Could not install hook"
        end if

    else
        msgBox "Hook is already enabled"
    end if

end sub ' }

sub storeKeyEvent(ev as keyEv) ' {
    curKeyEvent = curKeyEvent + 1
    if curKeyEvent >= nofKeyEventsStored then curKeyEvent = 0
    lastKeyEvents(curKeyEvent) = ev
end sub ' }

function getLastKeyEvent(n as byte) as keyEv ' {
    dim ix as byte

    if n > curKeyEvent then
       ix = nofKeyEventsStored - n
    else
       ix = curKeyEvent - n
    end if

    dim ev as keyEv
    ev = lastKeyEvents(ix)

    getLastKeyEvent = ev

end function ' }

sub initLastKeyevents() ' {

    dim cnt as byte
    dim ev  as keyEv

    ev.pressed = false
    ev.vkCode     = 0

    curKeyEvent = 0

    for cnt = 0 to nofKeyEventsStored - 1
        call storeKeyEvent(ev)
    next cnt
        

end sub ' }

public sub StopTaskAutomator() ' {

    if hookStarted then
       UnhookWindowsHookEx hookHandle
       hookStarted = false
    end if

    cells.clear

    debug.print "Tasks Automator finished"

end sub ' }

function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long ' {

'   dim upOrDown as string
'   dim altKey   as boolean
'   dim char     as string

    if nCode <> HC_ACTION then
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if

    if lParam.vkCode = VK_ESCAPE then StopTaskAutomator

'   select case wParam
'          case WM_KEYDOWN   : upOrDown = "keyDown"
'          case WM_KEYUP     : upOrDown = "keyUp"
'          case WM_SYSKEYDOWN: upOrDown = "sysDown"
'          case WM_SYSKEYUP  : upOrDown = "sysUp"
'   end select

    if wParam = WM_KEYDOWN or wParam = WM_SYSKEYDOWN then
       nextKeyEv.pressed = true
    else
       nextKeyEv.pressed = false
    end if

    nextKeyEv.vkCode = lParam.vkCode

    call storeKeyEvent(nextKeyEv)

'   dim ev0, ev2 as keyEv

    dim ev0 as keyEv
    dim ev2 as keyEv

    ev0 = getLastKeyEvent(0)
    ev2 = getLastKeyEvent(2)

    if ev0.pressed and ev0.vkCode = VK_RCONTROL and ev2.pressed and ev2.vkCode = VK_RCONTROL then
       call Beep(440, 200)
    end if


'  '
'  ' Apparently, the 5th bit is set if an ALT key was involved:
'  ' 
'    altKey = lParam.flags and 32
'
'    if ( lParam.vkCode >= asc("A") ) and ( lParam.vkCode <= asc("Z") ) then
'       char = chr(lParam.vkCode)
'
'    else
'
'      select case lParam.vkCode ' {
'             case VK_ESCAPE   : char = "esc"
'
'             case VK_LCONTROL : char = "ctrl L"
'             case VK_RCONTROL : char = "ctrl R"
'
'             case VK_LMENU    : char = "menu L"
'             case VK_RMENU    : char = "menu R"
'
'             case VK_RIGHT    : char = ">>"
'             case VK_LEFT     : char = "<<"
'             case VK_UP       : char = "^^"
'             case VK_DOWN     : char = "vv"
'
'             case VK_LSHIFT   : char = "shift L"
'             case VK_RSHIFT   : char = "shift R"
'
'             case VK_LWIN     : char = "win L"
'             case VK_RWIN     : char = "win R"
'
'             case VK_RETURN   : char = "enter"
'             case else        : char = "?"
'      end select ' }
'
'    end if
'
'  '
'  ' Display what the user has pressed.
'  '(Needs Excel)
'  '
'    cells(1,1) = upOrDown
'    cells(1,2) = lParam.vkCode
'    cells(1,3) = char
'    cells(1,4) = lParam.flags
'
'    if altKey then cells(1,5) = "alt" else cells(1,5) = "-"


    LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)

end function ' }
