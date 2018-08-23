option explicit

type keyEv ' {
     pressed    as boolean
     vkCode     as byte     ' VK_*
end type ' }

dim hookHandle  as long
dim hhShell     as long

dim   expectingCommand as boolean
dim   commandSoFar     as string

const nofKeyEventsStored = 20
dim   lastKeyEvents(nofKeyEventsStored) as keyEv
dim   curKeyEvent as byte
dim   nextKeyEv   as keyEv
dim   isSendingInput as boolean

dim   dbg as dbgFile


sub setHook(byRef hh as long, idHook as long, callBack as long) ' {

    if hh <> 0 then
       msgBox "Hook is already enabled"
       exit sub
    end if

    hh = SetWindowsHookEx(                 _
         idHook                          , _
         callBack                        , _
         GetModuleHandle(vbNullString)   , _
         0 )

    if hh = 0 then
       msgBox "could not install hook, " & GetLastError()
       exit sub
    end if

    debug.print "Hook started, hh = " & hh

end sub ' }

function unsetHook(byRef hh as long) as boolean ' {

    if hh <> 0 then
       UnhookWindowsHookEx hh
       hh = 0
       unsetHook = true
       exit function
     end if

     unsetHook = false
end function ' }

sub StartTaskAutomator() ' {

    call initLastKeyevents()
    nextKeyEv.pressed = false
    expectingCommand  = false
    isSendingInput    = false

    call setHook(hookHandle, WH_KEYBOARD_LL, addressOf LowLevelKeyboardProc)

  if dbg is nothing then
     set dbg = new dbgFile
     dbg.init("f:\task-automator\debug.out")
  end if

    dbg.text "TaskAutomator started"
end sub ' }

sub debugTA(txt as string) ' {
    selection.typeText txt
  selection.typeParagraph
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

    if unsetHook(hookHandle) then
       debug.print "Task Automator stopped"
    else
       debug.print "Task Automator was already stopped"
    end if

  if not dbg is nothing then
     set dbg = nothing
  end if

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

    if GetComputerName_ = "THINKPAD" then

 '    if isEventEqual(1, VK_RMENU   , true ) and _
 '       isEventEQual(0, VK_RMENU   , false) then
      if isEventEqual(1, VK_RCONTROL, true ) and _
         isEventEQual(0, VK_RCONTROL, false) then
            altGrPressed = true
       else
          altGrPressed   = false
       end if

    else

      if isEventEqual(3, VK_LCONTROL, true ) and _
         isEventEqual(2, VK_RMENU   , true ) and _
         isEventEqual(1, VK_LCONTROL, false) and _
         isEventEQual(0, VK_RMENU   , false) then
            altGrPressed = true
       else
          altGrPressed = false
       end if

    end if

end function ' }

function cmdInitSequence() as boolean ' {
    cmdInitSequence = altGrPressed
end function ' }

sub goToWindow(hWnd as long) ' {

    dbg.indent "goToWindow"

    dim curForegroundThreadId as long
    dim newForegroundThreadId as long

    if hWnd = 0 then
       debug.print "goToWindow, hWnd = 0"
       exit sub
    end if

    curForegroundThreadId = GetWindowThreadProcessId(GetForegroundWindow(), byVal 0&)
    newForegroundThreadID = GetWindowThreadProcessId(hWnd                 , byVal 0&)

    dim  rc as long
    call AttachThreadInput(curForegroundThreadId, newForegroundThreadID, true)
    rc = SetForeGroundWindow(hWnd)
    call AttachThreadInput(curForegroundThreadId, newForegroundThreadID, false)

'   if rc = 0 then
  dbg.text "SetForegroundWindow: rc = "  & rc
 '  debug.print "! Failed to SetForeGroundWindow"
 '  else
       if IsIconic(hWnd) then
          dbg.text "hWnd is iconic"
          call ShowWindow(hWnd, SW_RESTORE)
       else
          dbg.text "hWnd is not iconic"
          call ShowWindow(hWnd, SW_SHOW   )
       end if
 '  end if

    dbg.text "hWnd = " & hWnd

    call ShowWindow(hWnd, SW_SHOW)

  dbg.dedent

end sub ' }

sub SendInputText_TA(txt as string) ' {
    isSendingInput = true
    SendInputText    txt
    isSendingInput = false
end sub ' }

sub goToWindowWithClassAndCaption(class as string, caption as string) ' {
    dim hWnd as long
    hWnd = FindWindow(class, caption)
    goToWindow hWnd
end sub ' }

' sub goToWindowVBA() ' {
'     dim hWnd as long
'     hWnd = FindWindow("wndclass_desked_gsk", vbNullString)
'     goToWindow hWnd
' end sub ' }

function checkCommand(cmd as string) as boolean ' {

'   debug.print("Check Command " & cmd)

    if cmd = "STOP" then
       call stopTaskAutomator
       checkCommand = false
       exit function
    end if

    if cmd = "BLA" then ' {
       isSendingInput = true
       SendInputText "BlaBla01++\n"
       isSendingInput = false
       checkCommand = false
       exit function
    end if ' }

    if cmd = "INCA" then ' {
       isSendingInput = true
       SendInputText "Incident analysis\n"
       isSendingInput = false
       checkCommand = false
       exit function
    end if ' }

    if cmd = "CERT" then ' {
       dim r as RECT
       dim hWndSec as long
     ' hWndSec = FindWindow_ClassName_WindowText("#32770", "Windows Security")
       hWndSec = FindWindow("#32770", "Windows Security")
       r = GetWindowRect_(hWndSec)
       debug.print("hWndSec = " & hWndSec & ", left: " & r.left & ", top: " & r.top & ", right: " & r.right & ", bottom: " & r.bottom)

     ' call msgBox(r.left)
       checkCommand = false
       exit function
    end if ' }

    if cmd = "FG" then ' {
       debug.print "foreground Window is: " & GetForegroundWindow()
       debug.print "a: " & GetKeyboardLayout(0)
       debug.print "b: " & GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow(), 0))

       checkCommand = false
       exit function
    end if ' }

    if cmd = "EXCL" then ' {
       goToWindowWithClassAndCaption "XLMAIN", vbNullString
'      dim hWndExcel as long
'      hWndExcel = FindWindow("XLMAIN", vbNullString)
'    ' hWndExcel = FindWindow_ClassName("XLMAIN")
'      goToWindow hWndExcel
       checkCommand = false
       exit function
    end if ' }

    if cmd = "VBA" then ' {
       goToWindowWithClassAndCaption "wndclass_desked_gsk", vbNullString
       checkCommand = false
       exit function
    end if ' }

    if cmd = "NOTE" then ' {
       debugTA "NOTE"
       dim hWndNotepad as long
       hWndNotepad = FindWindow("Notepad++", vbNullString)

       if hWndNotepad <> 0 then
          debugTA "hWndNotepad <> 0 -> goToWindow hWndNotepad"
          goToWindow hWndNotepad
       else
          debug.print "Notepad window not found, opening it"
          shellOpen "C:\Users\RNyffenegger\AppData\Local\Microsoft\AppV\Client\Integration\F1272FE5-4D0A-4FA1-903F-AAE67B75C89E\Root\VFS\ProgramFilesX86\Notepad++\notepad++.exe"
       end if
     ' goToWindowWithClassAndCaption "Notepad", vbNullString

       checkCommand = false
       exit function
    end if ' }

    if cmd = "HELP" then ' {
       dim hWndHelpLine as long
       hWndHelpLine = FindWindow_WindowNameContains("ClassicDesk Prod 6.2")

       debug.print "hWndHelpLine = " & hWndHelpLine

       if hWndHelpLine = 0 then
          debug.print "hWndHelpLine = 0, trying to open executable"
          shellOpen "C:\ProgramData\Microsoft\AppV\Client\Integration\57DDD50C-7190-4CC4-9D25-365DC4F0E272\Root\VFS\ProgramFilesX86\helpLine\ClassicDesk.exe"
          checkCommand = false
          exit function
       end if

       goToWindow hWndHelpLine
       checkCommand = false
       exit function
    end if ' }

    if cmd = "CARK" then ' { Cyberark
       shellOpen "https://cyberark.wmhub.tq84.net/PasswordVault/auth/pki/"
       checkCommand = false
       exit function
    end if ' }

    if cmd = "SMCL" then ' { Avaloq's so called Smart Client
       shellOpen "L:\AvaloqSC\Avaloq_Test_DEV\SmartClientLauncher\SmartClientLauncher.exe", "contactcenter", "L:\AvaloqSC\Avaloq_Test_DEV\SmartClientLauncher"
    end if ' }

    if cmd = "SQLD" then ' { SQL Developer
       goToWindowWithClassAndCaption "SunAwtFrame", vbNullString
       checkCommand = false
       exit function
    end if ' }

    if cmd = "MVL" then ' { Move Window to left monitor
       dim hWndToMove as long
       hWndToMove = GetForegroundWindow
       MoveWindow HWndToMove, 1921, 0, 1920, 1080, true
    end if' }

    if cmd = "MVA" then ' { Move Window to right monitor
       dim hWndToMoveR as long
       hWndToMoveR = GetForegroundWindow
       MoveWindow HWndToMoveR, 0, 0, 1920, 1080, true
    end if' }


    if len(cmd) >= 4 then
       checkCommand = false
       exit function
    end if

  ' checkCommand = true
    checkCommand = checkTaskAutomatorCommand(cmd)

end function ' }

function ShellProc(byVal nCode as long, byVal wParam as long, lParam as long) ' {

    debug.print "ShellProc"

    if nCode = HSHELL_WINDOWCREATED then
       debug.print "a Windows was created"
    end if

    ShellProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)

end function ' }

function keyUpOrDown(wParam as long) as string ' {

    select case wParam
           case WM_KEYDOWN   : keyUpOrDown = "keyDown"
           case WM_KEYUP     : keyUpOrDown = "keyUp"
           case WM_SYSKEYDOWN: keyUpOrDown = "sysDown"
           case WM_SYSKEYUP  : keyUpOrDown = "sysUp"
    end select

end function ' }

function keyChar(vkCode as long) as string ' {

    if   ( ( vkCode >= asc("A") ) and ( vkCode <= asc("Z") ) ) or _
         ( ( vkCode >= asc("0") ) and ( vkCode <= asc("9") ) )    then

      keyChar = chr(vkCode)

    else

      select case vkCode ' {
             case VK_TAB      : keyChar = "tab"
             case VK_ESCAPE   : keyChar = "esc"
             case VK_PAUSE    : keyChar = "pause"
             case VK_PRIOR    : keyChar = "prior"
             case VK_NEXT     : keyChar = "next"

             case VK_LCONTROL : keyChar = "ctrl-L"
             case VK_RCONTROL : keyChar = "ctrl-R"

             case VK_LMENU    : keyChar = "menu-L"
             case VK_RMENU    : keyChar = "menu-R"

             case VK_RIGHT    : keyChar = ">>"
             case VK_LEFT     : keyChar = "<<"
             case VK_UP       : keyChar = "^^"
             case VK_DOWN     : keyChar = "vv"

             case VK_LSHIFT   : keyChar = "shift-L"
             case VK_RSHIFT   : keyChar = "shift-R"

             case VK_LWIN     : keyChar = "win-L"
             case VK_RWIN     : keyChar = "win-R"

             case VK_F1 to VK_F24 : keyChar = "F" & (24 - VK_F24 + vkCode)

             case VK_OEM_PLUS : keyChar = "OEM +"

             case VK_RETURN   : keyChar = "enter"
             case else        : keyChar = "0x" & hex(vkCode)
      end select ' }

'   if lParam.vkCode >= cLng("&h090") and lParam.vkCode <= cLng("&h0fc") then
'      debug.print "lParam.vkCode hex = " & hex(lParam.vkCode)
'   else
'      debug.print "lParam.vkCode chr = " & chr(lParam.vkCode)
'   end if

    end if

end function ' }

function LowLevelKeyboardProc(byVal nCode as Long, byVal wParam as long, lParam as KBDLLHOOKSTRUCT) as long ' {
'
'   Return value:
'     MSDN says: If the hook procedure processed the message, it may return a nonzero
'                value to prevent the system from passing the message to the rest of
'                the hook chain or the target window procedure.
'

    dim upOrDown as string
'   dim altKey   as boolean
    dim char     as string

    dim keyEventString as string

    if nCode <> HC_ACTION then
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if

    upOrDown = keyUpOrDown(wParam       )
    char     = keyChar    (lParam.vkCode)

    keyEventString = char & " " & upOrDown
  '
  ' Apparently, the 5th bit is set if an ALT key was involved:
  '
  ' altKey = lParam.flags and 32

'   debug.print char & " " & upOrDown

    if isSendingInput then
     ' debug.print keyEventString & " - Is sending input"
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if
  ' debug.print keyEventString

    if isEventEqual(0, VK_PAUSE, false) then
       StopTaskAutomator
    end if

'   if expectingCommand then
'      debug.print "expecting command"
'   else
'      debug.print "not expecting command"
'   end if



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

         ' 2018-07-25
         ' LowLevelKeyboardProc = 1
           LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
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

           LowLevelKeyboardProc = 1
           exit function

    end if
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

'   debug.print "Calling next LowLevelKeyboardProc"
    LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)

end function ' }
