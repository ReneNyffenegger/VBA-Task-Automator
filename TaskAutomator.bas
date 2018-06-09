option explicit

type keyEv ' {
     pressed    as boolean
     vkCode     as byte     ' VK_*
end type ' }

dim hookHandle  as long 
dim hookStarted as boolean

dim   expectingCommand as boolean
dim   lenCommand       as byte

const nofKeyEventsStored = 20
dim   lastKeyEvents(nofKeyEventsStored) as keyEv
dim   curKeyEvent as byte
dim   nextKeyEv   as keyEv


sub StartTaskAutomator() ' {

    call initLastKeyevents()
    nextKeyEv.pressed = false
    expectingCommand  = false


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

    dim ev0 as keyEv
    dim ev2 as keyEv

    ev0 = getLastKeyEvent(1)
    ev2 = getLastKeyEvent(3)

    if     ev0.pressed and ev0.vkCode = VK_RCONTROL and ev2.pressed and ev2.vkCode = VK_RCONTROL then

           call Beep(440, 200)
           expectingCommand = true
           lenCommand       = 0

    elseif expectingCommand then

           lenCommand = lenCommand + 1

           if lenCommand >= 4 then

              if chr(ev2.vkCode) = "E" and chr(ev0.vkCode) = "X" then
                call Beep(880, 200)

                dim hWndExcel as long

                  ' hWndExcel = FindWinow_WindowNameContains("neuer tab - google chrome")
                  ' hWndExcel = FindWinow_WindowNameContains("chrome")
                  ' hWndExcel = FindWinow_WindowNameContains("firefox")

                  ' hWndExcel = FindWinow_WindowNameContains("Excel")
                    hWndExcel = FindWinow_ClassName("XLMAIN")

                if hWndExcel = 0 then
                   debug.print "! Window not found"
                end if

                debug.print "hWnd = " & hWndExcel & ", parent: " & GetParent(hWndExcel) & ", parent parent: " & GetParent(GetParent(hWndExcel))


                dim curForegroundThreadId as long
                dim newForegroundThreadId as long

                    curForegroundThreadId = GetWindowThreadProcessId(GetForegroundWindow(), byVal 0&)
                    newForegroundThreadID = GetWindowThreadProcessId(hWndExcel            , byVal 0&)

                dim  rc as long
                call AttachThreadInput(curForegroundThreadId, newForegroundThreadID, true)
                rc = SetForeGroundWindow(hWndExcel)
                call AttachThreadInput(curForegroundThreadId, newForegroundThreadID, false)

                if rc = 0 then
                   debug.print "! Failed to SetForeGroundWindow"
                else
                   if IsIconic(hWndExcel) then
                      call ShowWindow(hWndExcel, SW_RESTORE)
                   else
                      call ShowWindow(hWndExcel, SW_SHOW   )
                   end if
                end if

'               debug.print "hWndExcel = " & hWndExcel

'               call ShowWindow(hWndExcel, SW_SHOW)

              end if

              expectingCommand = false
           else
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
