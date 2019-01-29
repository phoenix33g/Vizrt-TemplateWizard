

Sub updateLC(Sender)
    Dim lc: Set lc = findComponent("lc_" & Split(Sender.Name, "_")(1))
    Dim count: count = lc.Items.Count -1
    Dim pass: pass = Split(Sender.Name, "_")(2)
    Dim tempStr: tempStr = ""
    for i = 0 to count
        tempStr = UCase(lc.Items(i))
        if Sender.Checked then
           if InStr(tempStr, pass) <> 0 then lc.ItemIndex = i
        else
           if InStr(tempStr,"OUT") <> 0 then lc.ItemIndex = i
        end if
    next
End Sub
