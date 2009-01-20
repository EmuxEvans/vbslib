strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{impersonationLevel=impersonate}!\\" & _
       strComputer & "\root\cimv2")

For Each objclass in objWMIService.SubclassesOf()
    Wscript.Echo objClass.Path_.Class
Next
