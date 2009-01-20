strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")

For Each objclass in objWMIService.SubclassesOf()
    WScript.StdOut.Write objclass.Path_.Class
    arrDerivativeClasses = objClass.Derivation_ 
    For Each strDerivativeClass in arrDerivativeClasses 
       WScript.StdOut.Write " <- " & strDerivativeClass
    Next
    WScript.StdOut.Write vbNewLine
Next

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
