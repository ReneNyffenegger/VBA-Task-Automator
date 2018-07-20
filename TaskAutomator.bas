option explicit

type keyEv ' {
     pressed    as boolean
     vkCode     as byte     ' VK_*
end type ' }

dim hookHandle  as long
dim hookStarted as boolean

dim   expectingCommand as boolean
dim   commandSoFar     as string

const nofKeyEventsStored = 20
dim   lastKeyEvents(nofKeyEventsStored) as keyEv
dim   curKeyEvent as byte
dim   nextKeyEv   as keyEv


sub StartTaskAutomator() ' {

    call initLastKeyevents()
    nextKeyEv.pressed = false
    expectingCommand  = false


    if hookStarted = false then
        hookHandle = SetWindowsHookEx(       _
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

    debug.print "TaskAutomator started"
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

function isEventEqual(n as byte, vk as byte, pressed as boolean) as boolean ' {

    dim ev as keyEv
    ev  = getLastKeyEvent(n)
    if ev.pressed = pressed and ev.vkCode = vk  then
       isEventEqual = true
    else
       isEventEqual = false
    end if

end function ' }

function altGrPressed() as boolean ' {

    if isEventEqual(3, VK_LCONTROL, true ) and _
       isEventEqual(2, VK_RMENU   , true ) and _
       isEventEqual(1, VK_LCONTROL, false) and _
       isEventEQual(0, VK_RMENU   , false) then
          altGrPressed = true
    else
          altGrPressed = false
    end if

end function ' }

function cmdInitSequence() as boolean
    cmdInitSequence = altGrPressed
end function

sub goToWindow(hWnd as long) ' {

    dim curForegroundThreadId as long
    dim newForegroundThreadId as long

    curForegroundThreadId = GetWindowThreadProcessId(GetForegroundWindow(), byVal 0&)
    newForegroundThreadID = GetWindowThreadProcessId(hWnd            , byVal 0&)

    dim  rc as long
    call AttachThreadInput(curForegroundThreadId, newForegroundThreadID, true)
    rc = SetForeGroundWindow(hWnd)
    call AttachThreadInput(curForegroundThreadId, newForegroundThreadID, false)

    if rc = 0 then
       debug.print "! Failed to SetForeGroundWindow"
    else
       if IsIconic(hWnd) then
          call ShowWindow(hWnd, SW_RESTORE)
       else
          call ShowWindow(hWnd, SW_SHOW   )
       end if
    end if

    debug.print "hWnd = " & hWnd

    call ShowWindow(hWnd, SW_SHOW)

end sub ' }

function checkCommand(cmd as string) as boolean ' {

    debug.print("Check Command " & cmd)

    if cmd = "STOP" then
       call stopTaskAutomator
       checkCommand = false
       exit function
    end if

    if cmd = "EXCL" then
       dim hWndExcel as long
       hWndExcel = FindWinow_ClassName("XLMAIN")
       goToWindow hWndExcel
       checkCommand = false
       exit function
    end if

    if len(cmd) = 4 then
       checkCommand = false
       exit function
    end if

    checkCommand = true

end function ' }

function LowLevelKeyboardProc(byVal nCode as Long, ByVal wParam as Long, lParam as KBDLLHOOKSTRUCT) as long ' {

'   dim upOrDown as string
'   dim altKey   as boolean
'   dim char     as string

    if nCode <> HC_ACTION then
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if

    if lParam.vkCode = VK_ESCAPE then StopTaskAutomator

    if lParam.vkCode >= cLng("&h090") and lParam.vkCode <= cLng("&h0fc") then
       debug.print "lParam.vkCode = " & hex(lParam.vkCode)
    else
       debug.print chr(lParam.vkCode)
    end if

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


    if     cmdInitSequence then
           debug.print "starting new command"

           call Beep(440, 200)
           expectingCommand = true
           commandSoFar     = ""

           LowLevelKeyboardProc = 1
           exit function

    elseif expectingCommand then

           dim ev as keyEv
           dim c  as string

           ev = getLastKeyEvent(0)

           if not ev.pressed then
              if chr(ev.vkCode) >= "A" and chr(ev.vkCode) <= "Z" then
                  commandSoFar = commandSoFar + chr(ev.vkCode)
                  expectingCommand = checkCommand(commandSoFar)
              else
                 expectingCommand = false
              end if
           end if


           if expectingCommand then
              LowLevelKeyboardProc = 1
              exit function
           end if

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
