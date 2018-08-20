option explicit

function checkTaskAutomatorCommand(cmd as string) as boolean ' {

    if cmd = "FIX" then ' {
    '  isSendingInput = true
    '  SendInputText "Incident analysis\n"
    '  isSendingInput = false
       SendInputText_TA "Test evidence & fix script"
       checkTaskAutomatorCommand = false
       exit function
    end if ' }

    checkTaskAutomatorCommand = true
end function ' }
