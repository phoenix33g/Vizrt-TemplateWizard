'Created on 1/28/2019 for Univision SOTU banner
'Created by Israel Sanchez
'Returns an array of components that has a name that contains a sub-string given by the argument 'contains'

Function compArr(contains)
    Dim output : output = ""
    'Iterates through all components and finds component names that contain 'contains' 
    For i=0 to [_template].ComponentCount - 1
        Dim tempComp : Set tempComp = [_template].Components(i)
        If InStr(tempComp.name, contains) <> 0 then
           output = output & " " & tempComp.name
        End If
    Next
    output = Trim(output)
    output = Split(output)
    'Creates array from string of names.
    ReDim outputObj(UBound(output))
    For j=0 to UBound(output)
        Set outputObj(j) = findComponent(output(j))
    Next
    compArr = outputObj
End Function
