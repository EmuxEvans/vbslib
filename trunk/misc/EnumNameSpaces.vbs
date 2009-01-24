strComputer = "."
Call EnumNameSpaces("root")

Sub EnumNameSpaces(strNameSpace)
    On Error Resume Next
    WScript.Echo strNameSpace
    Set objWMIService=GetObject _
        ("winmgmts:{impersonationLevel=impersonate}\\" & _ 
            strComputer & "\" & strNameSpace)

    Set colNameSpaces = objWMIService.InstancesOf("__NAMESPACE")

    For Each objNameSpace in colNameSpaces
        Call EnumNameSpaces(strNameSpace & "\" & objNameSpace.Name)
    Next
End Sub
