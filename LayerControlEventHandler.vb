'Created on 1/28/2019 for Univision SOTU banner
'Created by Israel Sanchez
'An event handler built to work with a Transition Logic 'TVTWLayerControl'.
'This event handler could be placed on an 'TTWUniCheckBox' or 'TTWUniRadioButton' onClick event.
'This event handler uses naming conventions of components and toggle scene descriptions to link events.
'Example of component names: (TVTWLayerControl) lc_example; (TTWUniCheckBox) cb_example_IN, cb_example_A; (Scene Description) example_IN, example_OUT, example_A

Sub updateLC(Sender)
    Dim lc: Set lc = findComponent("lc_" & Split(Sender.Name, "_")(1)) 'Finds matching TVTWLayerControl <"lc_example">
    Dim count: count = lc.Items.Count -1
    Dim pass: pass = Split(Sender.Name, "_")(2)
    Dim tempStr: tempStr = ""    
    For i = 0 to count 'Iterates across(TVTWLayerControl) lc_example
        tempStr = UCase(lc.Items(i))
        if Sender.Checked then
           if InStr(tempStr, pass) <> 0 then lc.ItemIndex = i
        else
           if InStr(tempStr,"OUT") <> 0 then lc.ItemIndex = i
        end if
    Next
End Sub
