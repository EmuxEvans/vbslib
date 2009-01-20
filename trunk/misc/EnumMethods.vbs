strComputer = "."
strNameSpace = "root\cimv2"
strClass = "Win32_Service"

Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\" & strNameSpace & ":" & strClass)

WScript.Echo strClass & " Class Methods"
WScript.Echo "---------------------------"

For Each objClassMethod in objClass.Methods_
    WScript.Echo objClassMethod.Name
Next
  