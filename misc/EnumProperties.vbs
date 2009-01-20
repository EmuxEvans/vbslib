strComputer = "."
strNameSpace = "root\cimv2"
strClass = "Win32_Service"

Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\" & strNameSpace & ":" & strClass)

WScript.Echo strClass & " Class Properties"
WScript.Echo "------------------------------"

For Each objClassProperty in objClass.Properties_
    WScript.Echo objClassProperty.Name
Next
