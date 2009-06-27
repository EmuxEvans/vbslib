' VBScript Portable Library

Option Explicit

'============================================
'################ basic tool ################
'--------------------------------------------

Const RuntimeError = 51

Sub Bind(toStore, value)
  If IsObject(value) Then
    Set toStore = value
  Else
    toStore = value
  End If
End Sub

Sub BindAt(keyValueStore, key, value)
  If IsObject(value) Then
    Set keyValueStore(key) = value
  Else
    keyValueStore(key) = value
  End If
End Sub

Function Dictionary(keyValueList)
  Dim dict
  Set dict = CreateObject("Scripting.Dictionary")

  Dim isKey, key, i
  isKey = True

  For Each i In keyValueList
    If isKey Then
      Bind key, i
      dict.Add key, Empty
    Else
      BindAt dict, key, i
    End If
    isKey = Not isKey
  Next

  Set Dictionary = dict
End Function

' shortcut
Function D(keyValueList)
  Set D = Dictionary(keyValueList)
End Function

Function DictionaryMerge(dictionary1, dictionary2)
  Dim dict, key
  Set dict = CreateObject("Scripting.Dictionary")

  For Each key In dictionary1
    BindAt dict, key, dictionary1(key)
  Next

  For Each key In dictionary2
    BindAt dict, key, dictionary2(key)
  Next

  Set DictionaryMerge = dict
End Function

' shortcut
Function DMerge(dictionary1, dictionary2)
  Set DMerge = DictionaryMerge(dictionary1, dictionary2)
End Function

Function re(regexpPattern, regexpOptions)
  Dim regex, reOpts
  Set regex = New RegExp
  regex.Pattern = regexpPattern
  reOpts = LCase(regexpOptions)
  If InStr(reOpts, "i") > 0 Then
    regex.IgnoreCase = True
  End If
  If InStr(reOpts, "g") > 0 Then
    regex.Global = True
  End If
  If InStr(reOpts, "m") > 0 Then
    regex.Multiline = True
  End If
  Set re = regex
End Function

Class ListBuffer
  Private ivar_dict

  Private Sub Class_Initialize
    Set ivar_dict = CreateObject("Scripting.Dictionary")
  End Sub

  Public Property Get Count
    Count = ivar_dict.Count
  End Property

  Public Default Property Get Item(index)
    If ivar_dict.Exists(index) Then
      Bind Item, ivar_dict(index)
    Else
      Err.Raise 9, "stdlib.vbs:ListBuffer.Item(Get)", "out of range."
    End If
  End Property

  Public Property Let Item(index, value)
    If ivar_dict.Exists(index) Then
      ivar_dict(index) = value
    Else
      Err.Raise 9, "stdlib.vbs:ListBuffer.Item(Let)", "out of range."
    End If
  End Property

  Public Property Set Item(index, value)
    If ivar_dict.Exists(index) Then
      Set ivar_dict(index) = value
    Else
      Err.Raise 9, "stdlib.vbs:ListBuffer.Item(Set)", "out of range."
    End If
  End Property

  Public Property Get LastItem
    If ivar_dict.Count > 0 Then
      Bind LastItem, ivar_dict(ivar_dict.Count - 1)
    End If
  End Property

  Public Sub Add(value)
    Dim nextIndex
    nextIndex = ivar_dict.Count
    ivar_dict.Add nextIndex, value
  End Sub

  Public Sub Append(list)
    Dim i
    For Each i In list
      Add i
    Next
  End Sub

  Public Function Exists(index)
    Exists = ivar_dict.Exists(index)
  End Function

  Public Function Items
    ReDim itemList(ivar_dict.Count - 1)
    Dim i
    For i = 0 To ivar_dict.Count - 1
      BindAt itemList, i, ivar_dict(i)
    Next
    Items = itemList
  End Function

  Public Sub RemoveAll
    ivar_dict.RemoveAll
  End Sub

  Public Sub RemoveLastItem
    If ivar_dict.Count > 0 Then
      ivar_dict.Remove ivar_dict.Count - 1
    Else
      Err.Raise 9, "stdlib.vbs:ListBuffer.RemoveLastItem", "no item to remove."
    End If
  End Sub
End Class

Dim ShowString_Quote
Set ShowString_Quote = re("""", "g")

Function ShowString(value)
  ShowString = """" & ShowString_Quote.Replace(value, """""") & """"
End Function

Function ShowArray(value)
  Dim showList, i
  Set showList = New ListBuffer
  For Each i In value
    showList.Add ShowValue(i)
  Next
  ShowArray = "[" & Join(showList.Items, ",") & "]"
End Function

Function ShowDictionary(value)
  Dim showList, k
  Set showList = New ListBuffer
  For Each k In value.Keys
    showList.Add ShowValue(k) & "=>" & ShowValue(value(k))
  Next
  ShowDictionary = "{" & Join(showList.Items, ",") & "}"
End Function

Function ShowObject(value)
  Dim r
  Err.Clear
  On Error Resume Next
  r = ShowDictionary(value)
  If Err.Number <> 0 Then
    Err.Clear
    r = ShowArray(value)
  End If
  If Err.Number <> 0 Then
    Err.Clear
    r = ShowArray(value.Items)
  End If
  If Err.Number <> 0 Then
    Err.Clear
    r = "<" & TypeName(value) & ">"
  End If
  ShowObject = r
End Function

Function ShowOther(value)
  Dim r
  Err.Clear
  On Error Resume Next
  r = CStr(value)
  If Err.Number <> 0 Then
    Err.Clear
    r = ShowArray(value)
  End If
  If Err.Number <> 0 Then
    Err.Clear
    r = ShowDictionary(value)
  End If
  If Err.Number <> 0 Then
    Err.Clear
    r = ShowUnknown(value)
  End If
  ShowOther = r
End Function

Function ShowUnknown(value)
  ShowUnknown = "<unknown:" & VarType(value) & " " & TypeName(value) & ">"
End Function

Function ShowValue(value)
  Dim r
  Err.Clear
  On Error Resume Next

  If VarType(value) = vbString Then
    r = ShowString(value)
  ElseIf IsArray(value) Then
    r = ShowArray(value)
  ElseIf IsObject(value) Then
    r = ShowObject(value)
  ElseIf IsEmpty(value) Then
    r = "<empty>"
  ElseIf IsNull(value) Then
    r = "<null>"
  Else
    r = ShowOther(value)
  End If

  If Err.Number <> 0 Then
    Err.Clear
    r = ShowUnknown(value)
  End If

  ShowValue = r
End Function


'=================================================
'################ object accessor ################
'-------------------------------------------------

Dim ObjectProperty_AccessorPool
Set ObjectProperty_AccessorPool = CreateObject("Scripting.Dictionary")

Function ObjectProperty_CreateAccessor(name)
  Dim className, classExpr
  className = "ObjectProperty_Accessor_" & Name
  Set classExpr = New ListBuffer

  classExpr.Add "Class " & className
  classExpr.Add "  Public Default Property Get Item(obj)"
  classExpr.Add "    Bind Item, obj." & name
  classExpr.Add "  End Property"
  classExpr.Add ""
  classExpr.Add "  Public Property Let Item(obj, value)"
  classExpr.Add "    obj." & name & " = value"
  classExpr.Add "  End Property"
  classExpr.Add ""
  classExpr.Add "  Public Property Set Item(obj, value)"
  classExpr.Add "    Set obj." & name & " = value"
  classExpr.Add "  End Property"
  classExpr.Add "End Class"

  ExecuteGlobal Join(classExpr.Items, vbNewLine)
  Set ObjectProperty_CreateAccessor = Eval("New " & className)
End Function

Function ObjectProperty_GetAccessor(name)
  Dim key: key = UCase(name)
  If Not ObjectProperty_AccessorPool.Exists(key) Then
    ObjectProperty_AccessorPool.Add key, ObjectProperty_CreateAccessor(name)
  End If
  Set ObjectProperty_GetAccessor = ObjectProperty_AccessorPool(key)
End Function

Function GetObjectProperty(obj, name)
  Bind GetObjectProperty, ObjectProperty_GetAccessor(name)(obj)
End Function

Sub SetObjectProperty(obj, name, value)
  BindAt ObjectProperty_GetAccessor(name), obj, value
End Sub

Function ObjectPropertyExists(obj, Name)
  Err.Clear
  On Error Resume Next
  ObjectProperty_GetAccessor(name)(obj)
  Select Case Err.Number
    Case 0:
      Err.Clear
      ObjectPropertyExists = True
    Case 438:
      Err.Clear
      ObjectPropertyExists = False
    Case Else:
      Dim errNum, errSrc, errDsc
      errNum = Err.Number
      errSrc = Err.Source
      errDsc = Err.Description
      Err.Clear
      On Error GoTo 0
      Err.Raise errNum, errSrc, errDsc
  End Select
End Function


'===============================================
'################ object method ################
'-----------------------------------------------

Dim ObjectMethod_HandlerPool
Set ObjectMethod_HandlerPool = CreateObject("Scripting.Dictionary")

Function ObjectMethod_CreateHandler(name, argCount)
  Dim i, sep: sep = ""
  Dim argList: argList = ""
  For i = 0 To argCount - 1
    argList = argList & sep & "args(" & i & ")"
    sep = ", "
  Next

  Dim className, classExpr
  className = "ObjectMethod_Handler_" & name & "_Arg" & argCount
  Set classExpr = New ListBuffer

  classExpr.Add "Class " & className
  classExpr.Add "  Public Sub ExecSubProc(obj, args)"
  classExpr.Add "    obj." & name & " " & argList
  classExpr.Add "  End Sub"
  classExpr.Add ""
  classExpr.Add "  Public Function ExecFuncProc(obj, args)"
  classExpr.Add "    Bind ExecFuncProc, obj." & name & "(" & argList & ")"
  classExpr.Add "  End Function"
  classExpr.Add "End Class"

  ExecuteGlobal Join(classExpr.Items, vbNewLine)
  Set ObjectMethod_CreateHandler = Eval("New " & className)
End Function

Function ObjectMethod_GetHandler(name, argCount)
  Dim key: key = UCase(name) & "_" & argCount
  If Not ObjectMethod_HandlerPool.Exists(key) Then
    ObjectMethod_HandlerPool.Add key, ObjectMethod_CreateHandler(name, argCount)
  End If
  Set ObjectMethod_GetHandler = ObjectMethod_HandlerPool(key)
End Function

Sub ExecObjectMethodSubProc(obj, name, args)
  Dim argCount, method
  argCount = CountItem(args)
  Set method = ObjectMethod_GetHandler(name, argCount)
  method.ExecSubProc obj, args
End Sub

Function ExecObjectMethodFuncProc(obj, name, args)
  Dim argCount, method
  argCount = CountItem(args)
  Set method = ObjectMethod_GetHandler(name, argCount)
  Bind ExecObjectMethodFuncProc, method.ExecFuncProc(obj, args)
End Function

Dim ObjectMethod_ProcBuilderPool
Set ObjectMethod_ProcBuilderPool = CreateObject("Scripting.Dictionary")

Function ObjectMethod_CreateProcBuilder(name, argCount)
  Dim i, sep: sep = ""
  Dim argList: argList = ""
  For i = 1 To argCount
    argList = argList & sep & "arg" & i
    sep = ", "
  Next

  Dim className, classExpr
  className = "ObjectMethod_Proc_" & name & "_Arg" & argCount
  Set classExpr = New ListBuffer

  classExpr.Add "Class " & className & "_SubProc"
  classExpr.Add "  Private ivar_obj"
  classExpr.Add ""
  classExpr.Add "  Public Property Set Self(obj)"
  classExpr.Add "    Set ivar_obj = obj"
  classExpr.Add "  End Property"
  classExpr.Add ""
  classExpr.Add "  Public Default Sub Execute(" & argList & ")"
  classExpr.Add "    ivar_obj." & name & " " & argList
  classExpr.Add "  End Sub"
  classExpr.Add "End Class"
  classExpr.Add ""
  classExpr.Add "Class " & className & "_FuncProc"
  classExpr.Add "  Private ivar_obj"
  classExpr.Add ""
  classExpr.Add "  Public Property Set Self(obj)"
  classExpr.Add "    Set ivar_obj = obj"
  classExpr.Add "  End Property"
  classExpr.Add ""
  classExpr.Add "  Public Default Function Execute(" & argList & ")"
  classExpr.Add "    Bind Execute, ivar_obj." & name & "(" & argList & ")"
  classExpr.Add "  End Function"
  classExpr.Add "End Class"
  classExpr.Add ""
  classExpr.Add "Class " & className & "_Builder"
  classExpr.Add "  Public Function CreateSubProc(obj)"
  classExpr.Add "    Dim proc"
  classExpr.Add "    Set proc = New " & className & "_SubProc"
  classExpr.Add "    Set proc.Self = obj"
  classExpr.Add "    Set CreateSubProc = proc"
  classExpr.Add "  End Function"
  classExpr.Add ""
  classExpr.Add "  Public Function CreateFuncProc(obj)"
  classExpr.Add "    Dim proc"
  classExpr.Add "    Set proc = New " & className & "_FuncProc"
  classExpr.Add "    Set proc.Self = obj"
  classExpr.Add "    Set CreateFuncProc = proc"
  classExpr.Add "  End Function"
  classExpr.Add "End Class"

  ExecuteGlobal Join(classExpr.Items, vbNewLine)
  Set ObjectMethod_CreateProcBuilder = Eval("New " & className & "_Builder")
End Function

Function ObjectMethod_GetProcBuilder(name, argCount)
  Dim key: key = UCase(name) & "_" & argCount
  If Not ObjectMethod_ProcBuilderPool.Exists(key) Then
    ObjectMethod_ProcBuilderPool.Add key, ObjectMethod_CreateProcBuilder(name, argCount)
  End If
  Set ObjectMethod_GetProcBuilder = ObjectMethod_ProcBuilderPool(key)
End Function

Function GetObjectMethodSubProc(obj, name, argCount)
  Set GetObjectMethodSubProc = ObjectMethod_GetProcBuilder(name, argCount).CreateSubProc(obj)
End Function

Function GetObjectMethodFuncProc(obj, name, argCount)
  Set GetObjectMethodFuncProc = ObjectMethod_GetProcBuilder(name, argCount).CreateFuncProc(obj)
End Function


'==================================================
'################ procedure subset ################
'--------------------------------------------------

Dim ProcSubset_ProcBuilderPool
Set ProcSubset_ProcBuilderPool = CreateObject("Scripting.Dictionary")

Function ProcSubset_IndexExists(paramIndexList, index)
  Dim i
  For Each i In paramIndexList
    If i = index Then
      ProcSubset_IndexExists = True
      Exit Function
    End If
  Next
  ProcSubset_IndexExists = False
End Function

Function ProcSubset_CreateProcBuilder(argCount, paramIndexList)
  Dim className, classExpr
  className = "ProcSubset_Arg" & argCount & "_" & Join(paramIndexList, "_")
  Set classExpr = New ListBuffer

  Dim argList, applyArgList, i
  Set argList = New ListBuffer
  Set applyArgList = New ListBuffer
  For i = 0 To argCount - 1
    If ProcSubset_IndexExists(paramIndexList, i) Then
      applyArgList.Add "ivar_arg" & i
    Else
      argList.Add "arg" & i
      applyArgList.Add "arg" & i
    End If
  Next
  argList = Join(argList.Items, ", ")
  applyArgList = Join(applyArgList.Items, ", ")

  classExpr.Add "Class " & className & "_SubProc"
  classExpr.Add "  Private ivar_proc"
  For Each i In paramIndexList
    classExpr.Add "  Private ivar_arg" & i
  Next
  classExpr.Add ""
  classExpr.Add "  Public Property Set Proc(value)"
  classExpr.Add "    Set ivar_proc = value"
  classExpr.Add "  End Property"
  classExpr.Add ""
  For Each i In paramIndexList
    classExpr.Add "  Public Property Let Arg" & i & "(value)"
    classExpr.Add "    ivar_arg" & i & " = value"
    classExpr.Add "  End Property"
    classExpr.Add ""
    classExpr.Add "  Public Property Set Arg" & i & "(value)"
    classExpr.Add "    Set ivar_arg" & i & " = value"
    classExpr.Add "  End Property"
    classExpr.Add ""
  Next
  classExpr.Add "  Public Default Sub Execute(" & argList & ")"
  classExpr.Add "    Call ivar_proc(" & applyArgList & ")"
  classExpr.Add "  End Sub"
  classExpr.Add "End Class"
  classExpr.Add ""
  classExpr.Add "Class " & className & "_FuncProc"
  classExpr.Add "  Private ivar_proc"
  For Each i In paramIndexList
    classExpr.Add "  Private ivar_arg" & i
  Next
  classExpr.Add ""
  classExpr.Add "  Public Property Set Proc(value)"
  classExpr.Add "    Set ivar_proc = value"
  classExpr.Add "  End Property"
  classExpr.Add ""
  For Each i In paramIndexList
    classExpr.Add "  Public Property Let Arg" & i & "(value)"
    classExpr.Add "    ivar_arg" & i & " = value"
    classExpr.Add "  End Property"
    classExpr.Add ""
    classExpr.Add "  Public Property Set Arg" & i & "(value)"
    classExpr.Add "    Set ivar_arg" & i & " = value"
    classExpr.Add "  End Property"
    classExpr.Add ""
  Next
  classExpr.Add "  Public Default Function Execute(" & argList & ")"
  classExpr.Add "    Bind Execute, ivar_proc(" & applyArgList & ")"
  classExpr.Add "  End Function"
  classExpr.Add "End Class"
  classExpr.Add ""
  classExpr.Add "Class " & className & "_Builder"
  classExpr.Add "  Public Function CreateSubProc(proc)"
  classExpr.Add "    Dim subset"
  classExpr.Add "    Set subset = New " & className & "_SubProc"
  classExpr.Add "    Set subset.Proc = proc"
  classExpr.Add "    Set CreateSubProc = subset"
  classExpr.Add "  End Function"
  classExpr.Add ""
  classExpr.Add "  Public Function CreateFuncProc(proc)"
  classExpr.Add "    Dim subset"
  classExpr.Add "    Set subset = New " & className & "_FuncProc"
  classExpr.Add "    Set subset.Proc = proc"
  classExpr.Add "    Set CreateFuncProc = subset"
  classExpr.Add "  End Function"
  classExpr.Add "End Class"

  ExecuteGlobal Join(classExpr.Items, vbNewLine)
  Set ProcSubset_CreateProcBuilder = Eval("New " & className & "_Builder")
End Function

Function ProcSubset_GetProcBuilder(argCount, paramIndexList)
  Dim key
  key = "arg" & argCount & "_" & Join(paramIndexList, "_")
  If Not ProcSubset_ProcBuilderPool.Exists(key) Then
    ProcSubset_ProcBuilderPool.Add key, ProcSubset_CreateProcBuilder(argCount, paramIndexList)
  End If
  Set ProcSubset_GetProcBuilder = ProcSubset_ProcBuilderPool(key)
End Function

Dim ProcSubset_NumericCompare
Set ProcSubset_NumericCompare = GetRef("NumericCompare")

Function ProcSubset_BuildParamsPair(argCount, params)
  Dim paramIndexList, paramDict, i
  Set paramIndexList = New ListBuffer

  If IsArray(params) Then
    Set paramDict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(params)
      paramIndexList.Add i
      paramDict.Add i, params(i)
    Next
  Else
    Set paramDict = params
    For Each i In paramDict.Keys
      paramIndexList.Add i
    Next
  End If

  paramIndexList = paramIndexList.Items
  Sort paramIndexList, ProcSubset_NumericCompare

  ProcSubset_BuildParamsPair = Array(paramIndexList, paramDict)
End Function

Sub ProcSubset_SetParams(subset, paramDict)
  Dim i
  For Each i In paramDict.Keys
    SetObjectProperty subset, "arg" & i, paramDict(i)
  Next
End Sub

Function GetSubProcSubset(proc, argCount, params)
  Dim paramsPair
  paramsPair = ProcSubset_BuildParamsPair(argCount, params)

  Dim builder, subset
  Set builder = ProcSubset_GetProcBuilder(argCount, paramsPair(0))
  Set subset = builder.CreateSubProc(proc)
  ProcSubset_SetParams subset, paramsPair(1)

  Set GetSubProcSubset = subset
End Function

Function GetFuncProcSubset(proc, argCount, params)
  Dim paramsPair
  paramsPair = ProcSubset_BuildParamsPair(argCount, params)

  Dim builder, subset
  Set builder = ProcSubset_GetProcBuilder(argCount, paramsPair(0))
  Set subset = builder.CreateFuncProc(proc)
  ProcSubset_SetParams subset, paramsPair(1)

  Set GetFuncProcSubset = subset
End Function


'===============================================
'################ pseudo object ################
'-----------------------------------------------

Sub PseudoObject_AttachMethodSubProc(pseudoObject, key, proc, argCount)
  Set pseudoObject(key) = GetSubProcSubset(proc, argCount, Array(pseudoObject))
End Sub

Sub PseudoObject_AttachMethodFuncProc(pseudoObject, key, proc, argCount)
  Set pseudoObject(key) = GetFuncProcSubset(proc, argCount, Array(pseudoObject))
End Sub


'===========================================
'################ list tool ################
'-------------------------------------------

Function FirstItem(list)
  Dim i
  For Each i In list
    Bind FirstItem, i
    Exit Function
  Next
  Err.Raise 9, "stdlib.vbs:FirstItem", "empty list."
End Function

Function LastItem(list)
  Dim found, i
  found = False
  For Each i In list
    Bind LastItem, i
    found = True
  Next
  If Not found Then
    Err.Raise 9, "stdlib.vbs:LastItem", "empty list."
  End If
End Function

Function CountItem(list)
  If IsArray(list) Then
    CountItem = UBound(list) + 1
    Exit Function
  End If
  
  If IsObject(list) Then
    If ObjectPropertyExists(list, "Count") Then
      CountItem = list.Count
      Exit Function
    End If
  End If

  Dim count, i
  count = 0
  For Each i In list
    count = count + 1
  Next
  CountItem = count
End Function

Function Find(list, defaultValue, cond)
  Dim i
  Bind Find, defaultValue
  For Each i In list
    If cond(i) Then
      Bind Find, i
      Exit Function
    End If
  Next
End Function

Function FindPos(list, cond)
  Dim pos, i
  pos = 0
  For Each i In list
    If cond(i) Then
      FindPos = pos
      Exit Function
    End If
    pos = pos + 1
  Next
End Function

Function FindAll(list, cond)
  Dim findList, i
  Set findList = New ListBuffer
  For Each i In list
    If cond(i) Then
      findList.Add i
    End If
  Next
  FindAll = findList.Items
End Function

Function Map(list, func)
  Dim resultList, i
  Set resultList = New ListBuffer
  For Each i In list
    resultList.Add func(i)
  Next
  Map = resultList.Items
End Function

Function Inject(list, initialValue, func)
  Dim isFirst, result
  If IsEmpty(initialValue) Then
    isFirst = True
  Else
    isFirst = False
    Bind result, initialValue
  End If

  Dim i
  For Each i In list
    If isFirst Then
      Bind result, i
      isFirst = False
    Else
      Bind result, func(result, i)
    End If
  Next

  Bind Inject, result
End Function

Function Max(list, compare)
  Bind Max, Inject(list, Empty, GetFuncProcSubset(GetRef("PriorMax"), 3, Array(compare)))
End Function

Function Min(list, compare)
  Bind Min, Inject(list, Empty, GetFuncProcSubset(GetRef("PriorMin"), 3, Array(compare)))
End Function

Function Range(first, cond, increment)
  Dim resultList
  Set resultList = New ListBuffer

  Dim i
  Bind i, first

  Do While cond(i)
    resultList.Add i
    Bind i, increment(i)
  Loop

  Range = resultList.Items
End Function

Function NumericRange(first, last, incrementStep, excludeLast)
  Dim cond
  If excludeLast Then
    Set cond = ValueLessThan(last)
  Else
    Set cond = ValueLessEqual(last)
  End If

  NumericRange = Range(first, cond, _
                       GetFuncProcSubset(GetRef("Add"), 2, _
                                         Array(incrementStep)))
End Function

Function Numbering(first, last)
  Numbering = NumericRange(first, last, 1, False)
End Function

Function AryReverse(ByVal list)
  Dim l, r, t
  l = 0
  r = UBound(list)

  Do While l < r
    Bind t, list(l)
    BindAt list, l, list(r)
    BindAt list, r, t
    l = l + 1
    r = r - 1
  Loop

  AryReverse = list
End Function


'==================================================
'################ utility function ################
'--------------------------------------------------

Sub UtilityFunction_DefineVBScriptFunctionAliases
  Dim aliases
  Set aliases = New ListBuffer

  ' Data Type
  aliases.Add Array("CBool", 1)
  aliases.Add Array("CByte", 1)
  aliases.Add Array("CCur", 1)
  aliases.Add Array("CDate", 1)
  aliases.Add Array("CDbl", 1)
  aliases.Add Array("CInt", 1)
  aliases.Add Array("CLng", 1)
  aliases.Add Array("CSng", 1)
  aliases.Add Array("CStr", 1)
  aliases.Add Array("IsArray", 1)
  aliases.Add Array("IsDate", 1)
  aliases.Add Array("IsEmpty", 1)
  aliases.Add Array("IsNull", 1)
  aliases.Add Array("IsNumeric", 1)
  aliases.Add Array("IsObject", 1)
  aliases.Add Array("TypeName", 1)
  aliases.Add Array("VarType", 1)
  aliases.Add Array("LBound", "LBound", 1)
  aliases.Add Array("LBound", "LBound1", 1)
  aliases.Add Array("LBound", "LBound2", 2)
  aliases.Add Array("UBound", "UBound", 1)
  aliases.Add Array("UBound", "UBound1", 1)
  aliases.Add Array("UBound", "UBound2", 2)

  ' String
  aliases.Add Array("Asc", 1)
  aliases.Add Array("Chr", 1)
  aliases.Add Array("Len", 1)
  aliases.Add Array("LCase", 1)
  aliases.Add Array("UCase", 1)
  aliases.Add Array("Trim", 1)
  aliases.Add Array("LTrim", 1)
  aliases.Add Array("RTrim", 1)
  aliases.Add Array("Space", 1)
  aliases.Add Array("StrReverse", 1)
  aliases.Add Array("Join", 2)
  aliases.Add Array("Left", 2)
  aliases.Add Array("LeftB", 2)
  aliases.Add Array("Right", 2)
  aliases.Add Array("RightB", 2)
  aliases.Add Array("String", 2)
  aliases.Add Array("InStr", "InStr", 2)
  aliases.Add Array("InStr", "InStr2", 2)
  aliases.Add Array("InStr", "InStr3", 3)
  aliases.Add Array("InStr", "InStr4", 4)
  aliases.Add Array("InStrRev", "InStrRev", 2)
  aliases.Add Array("InStrRev", "InStrRev2", 2)
  aliases.Add Array("InStrRev", "InStrRev3", 3)
  aliases.Add Array("InStrRev", "InStrRev4", 4)
  aliases.Add Array("Mid", "Mid", 2)
  aliases.Add Array("Mid", "Mid2", 2)
  aliases.Add Array("Mid", "Mid3", 3)
  aliases.Add Array("MidB", "MidB", 2)
  aliases.Add Array("MidB", "MidB2", 2)
  aliases.Add Array("MidB", "MidB3", 3)
  aliases.Add Array("Replace", "Replace", 3)
  aliases.Add Array("Replace", "Replace3", 3)
  aliases.Add Array("Replace", "Replace4", 4)
  aliases.Add Array("Replace", "Replace5", 5)
  aliases.Add Array("Replace", "Replace6", 6)
  aliases.Add Array("Split", "Split", 1)
  aliases.Add Array("Split", "Split1", 1)
  aliases.Add Array("Split", "Split2", 2)
  aliases.Add Array("Split", "Split3", 3)
  aliases.Add Array("Split", "Split4", 4)
  aliases.Add Array("StrComp", "StrComp", 2)
  aliases.Add Array("StrComp", "StrComp2", 2)
  aliases.Add Array("StrComp", "StrComp3", 3)

  ' Number
  aliases.Add Array("Hex", 1)
  aliases.Add Array("Oct", 1)
  aliases.Add Array("Sgn", 1)
  aliases.Add Array("Int", 1)
  aliases.Add Array("Fix", 1)
  aliases.Add Array("Rnd", 1)
  aliases.Add Array("Round", "Round", 1)
  aliases.Add Array("Round", "Round1", 1)
  aliases.Add Array("Round", "Round2", 2)

  ' Math
  aliases.Add Array("Abs", 1)
  aliases.Add Array("Atan", 1)
  aliases.Add Array("Cos", 1)
  aliases.Add Array("Exp", 1)
  aliases.Add Array("Log", 1)
  aliases.Add Array("Sin", 1)
  aliases.Add Array("Sqr", 1)
  aliases.Add Array("Tan", 1)

  ' DateTime
  aliases.Add Array("DateValue", 1)
  aliases.Add Array("TimeValue", 1)
  aliases.Add Array("Year", 1)
  aliases.Add Array("Day", 1)
  aliases.Add Array("Month", 1)
  aliases.Add Array("Hour", 1)
  aliases.Add Array("Minute", 1)
  aliases.Add Array("Second", 1)
  aliases.Add Array("DateSerial", 3)
  aliases.Add Array("TimeSerial", 3)
  aliases.Add Array("DateAdd", 3)
  aliases.Add Array("WeekdayName", 3)
  aliases.Add Array("DateDiff", "DateDiff", 3)
  aliases.Add Array("DateDiff", "DateDiff3", 3)
  aliases.Add Array("DateDiff", "DateDiff4", 4)
  aliases.Add Array("DateDiff", "DateDiff5", 5)
  aliases.Add Array("DatePart", "DatePart", 2)
  aliases.Add Array("DatePart", "DatePart2", 2)
  aliases.Add Array("DatePart", "DatePart3", 3)
  aliases.Add Array("DatePart", "DatePart4", 4)
  aliases.Add Array("MonthName", "MonthName", 1)
  aliases.Add Array("MonthName", "MonthName1", 1)
  aliases.Add Array("MonthName", "MonthName2", 2)
  aliases.Add Array("Weekday", "Weekday", 1)
  aliases.Add Array("Weekday", "Weekday1", 1)
  aliases.Add Array("Weekday", "Weekday2", 2)

  ' Eval
  aliases.Add Array("Eval", 1)
  aliases.Add Array("GetRef", 1)

  ' Object
  aliases.Add Array("CreateObject", "CreateObject", 1)
  aliases.Add Array("CreateObject", "CreateObject1", 1)
  aliases.Add Array("CreateObject", "CreateObject2", 2)
  aliases.Add Array("GetObject", "GetObject", 1)
  aliases.Add Array("GetObject", "GetObject2", 2)

  Dim aliasDef, aliasExpr, name, alias, argCount, argList, sep, i
  For Each aliasDef In aliases.Items
    Select Case CountItem(aliasDef)
      Case 2:
        name = aliasDef(0)
        alias = aliasDef(0)
        argCount = aliasDef(1)
      Case 3:
        name = aliasDef(0)
        alias = aliasDef(1)
        argCount = aliasDef(2)
      Case Else:
        Err.Raise RuntimeError, _
           "stdlib.vbs:UtilityFunction_DefineVBScriptFunctionAliases", _
           "Unknown aliasDef: " & ShowValue(aliasDef)
    End Select

    argList = ""
    sep = ""
    For i = 0 To argCount - 1
      argList = argList & sep & "arg" & i
      sep = ", "
    Next

    Set aliasExpr = New ListBuffer
    aliasExpr.Add "Function " & alias & "_(" & argList & ")"
    aliasExpr.Add "  Bind " & alias & "_, " & name & "(" & argList & ")"
    aliasExpr.Add "End Function"

    ExecuteGlobal Join(aliasExpr.Items, vbNewLine)
  Next
End Sub
UtilityFunction_DefineVBScriptFunctionAliases

Function Equal(expected, value)
  Equal = (value = expected)
End Function

Function ValueEqual(expected)
  Set ValueEqual = _
      GetFuncProcSubset(GetRef("Equal"), 2, Array(expected))
End Function

Function GreaterThan(lowerBound, value)
  GreaterThan = value > lowerBound
End Function

Function ValueGreaterThan(lowerBound)
  Set ValueGreaterThan = _
      GetFuncProcSubset(GetRef("GreaterThan"), 2, Array(lowerBound))
End Function

Function GreaterEqual(lowerBound, value)
  GreaterEqual = value >= lowerBound
End Function

Function ValueGreaterEqual(lowerBound)
  Set ValueGreaterEqual = _
      GetFuncProcSubset(GetRef("GreaterEqual"), 2, Array(lowerBound))
End Function

Function LessThan(upperBound, value)
  LessThan = value < upperBound
End Function

Function ValueLessThan(upperBound)
  Set ValueLessThan = _
      GetFuncProcSubset(GetRef("LessThan"), 2, Array(upperBound))
End Function

Function LessEqual(upperBound, value)
  LessEqual = value <= upperBound
End Function

Function ValueLessEqual(upperBound)
  Set ValueLessEqual = _
      GetFuncProcSubset(GetRef("LessEqual"), 2, Array(upperBound))
End Function

Function Between(lowerBound, upperBound, value)
  Between = (lowerBound <= value) And (value <= upperBound)
End Function

Function BetweenExcludeUpperBound(lowerBound, upperBound, value)
  BetweenExcludeUpperBound = (lowerBound <= value) And (value < upperBound)
End Function

Function ValueBetween(lowerBound, upperBound, exclude)
  If exclude Then
    Set ValueBetween = _
        GetFuncProcSubset(GetRef("BetweenExcludeUpperBound"), 3, Array(lowerBound, upperBound))
  Else
    Set ValueBetween = _
        GetFuncProcSubset(GetRef("Between"), 3, Array(lowerBound, upperBound))
  End If
End Function

Function ValueMatch(regex)
  Set ValueMatch = GetObjectMethodFuncProc(regex, "Test", 1)
End Function

Function ValueFilterFunc(filter, func, value)
  Bind ValueFilterFunc, func(filter(value))
End Function

Function ValueFilter(filter, func)
  Set ValueFilter = GetFuncProcSubset(GetRef("ValueFilterFunc"), 3, Array(filter, func))
End Function

Function NotCondFunc(cond, value)
  NotCondFunc = Not cond(value)
End Function

Function NotCond(cond)
  Set NotCond = GetFuncProcSubset(GetRef("NotCondFunc"), 2, Array(cond))
End Function

Function AndCondFunc(cond1, cond2, value)
  AndCondFunc = cond1(value) And cond2(value)
End Function

Function AndCond(cond1, cond2)
  Set AndCond = GetFuncProcSubset(GetRef("AndCondFunc"), 3, Array(cond1, cond2))
End Function

Function OrCondFunc(cond1, cond2, value)
  OrCondFunc = cond1(value) Or cond2(value)
End Function

Function OrCond(cond1, cond2)
  Set OrCond = GetFuncProcSubset(GetRef("OrCondFunc"), 3, Array(cond1, cond2))
End Function

Function NumericCompare(a, b)
  NumericCompare = a - b
End Function

Function ReverseCompareFunc(compare, a, b)
  ReverseCompareFunc = compare(b, a)
End Function

Function ReverseCompare(compare)
  Set ReverseCompare = _
      GetFuncProcSubset(GetRef("ReverseCompareFunc"), 3, Array(compare))
End Function

Function CompareFilterFunc(filter, compare, a, b)
  CompareFilterFunc = compare(filter(a), filter(b))
End Function

Function CompareFilter(filter, compare)
  Set CompareFilter = _
      GetFuncProcSubset(GetRef("CompareFilterFunc"), 4, Array(filter, compare))
End Function

Function ObjectPropertyCompare(propertyName, propertyCompare)
  Set ObjectPropertyCompare = _
      CompareFilter(ValueObjectProperty(propertyName), propertyCompare)
End Function

Function CompareEqual(compare, expected, value)
  CompareEqual = (compare(value, expected) = 0)
End Function

Function CompareGreaterThan(compare, lowerBound, value)
  CompareGreaterThan = (compare(value, lowerBound) > 0)
End Function

Function CompareGreaterEqual(compare, lowerBound, value)
  CompareGreaterEqual = (compare(value, lowerBound) >= 0)
End Function

Function CompareLessThan(compare, upperBound, value)
  CompareLessThan = (compare(value, upperBound) < 0)
End Function

Function CompareLessEqual(compare, upperBound, value)
  CompareLessEqual = (compare(value, upperBound) <= 0)
End Function

Function ValueCompare(operatorType, bound, compare)
  Select Case operatorType
    Case "=":
      Set ValueCompare = _
          GetFuncProcSubset(GetRef("CompareEqual"), 3, Array(compare, bound))
    Case ">":
      Set ValueCompare = _
          GetFuncProcSubset(GetRef("CompareGreaterThan"), 3, Array(compare, bound))
    Case ">=":
      Set ValueCompare = _
          GetFuncProcSubset(GetRef("CompareGreaterEqual"), 3, Array(compare, bound))
    Case "<":
      Set ValueCompare = _
          GetFuncProcSubset(GetRef("CompareLessThan"), 3, Array(compare, bound))
    Case "<=":
      Set ValueCompare = _
          GetFuncProcSubset(GetRef("CompareLessEqual"), 3, Array(compare, bound))
    Case Else:
      Err.Raise 5, "stdlib.vbs:ValueCompare", "unknown operatorType: " & operatorType
  End Select
End Function

Function ValueReplace(regex, replace)
  Set ValueReplace = _
      GetFuncProcSubset(GetObjectMethodFuncProc(regex, "Replace", 2), 2, D(Array(1, replace)))
End Function

Function GetItemAt(keyValueStore, key)
  Bind GetItemAt, keyValueStore(key)
End Function

Function CollectItems(keyValueStore, keyValueGet, keyList)
  Dim dict
  Set dict = CreateObject("Scripting.Dictionary")

  Dim key
  For Each key In keyList
    BindAt dict, key, keyValueGet(keyValueStore, key)
  Next

  Set CollectItems = dict
End Function

Function ValueItemAt(key)
  Set ValueItemAt = GetFuncProcSubset(GetRef("GetItemAt"), 2, D(Array(1, key)))
End Function

Function ValueItemsAt(keyList)
  Set ValueItemsAt = _
      GetFuncProcSubset(GetRef("CollectItems"), 3, _
                        D(Array(1, GetRef("GetItemAt"), 2, keyList)))
End Function

Function ValueObjectProperty(propertyName)
  Set ValueObjectProperty = _
      GetFuncProcSubset(GetRef("GetObjectProperty"), 2, D(Array(1, propertyName)))
End Function

Function ValueObjectProperties(propertyNameList)
  Set ValueObjectProperties = _
      GetFuncProcSubset(GetRef("CollectItems"), 3, _
                        D(Array(1, GetRef("GetObjectProperty"), 2, propertyNameList)))
End Function

Function PriorMax(compare, a, b)
  If compare(a, b) > 0 Then
    Bind PriorMax, a
  Else
    Bind PriorMax, b
  End If
End Function

Function PriorMin(compare, a, b)
  If compare(a, b) < 0 Then
    Bind PriorMin, a
  Else
    Bind PriorMin, b
  End If
End Function

Function Add(number1, number2)
  Add = number1 + number2
End Function

Function Subtract(number1, number2)
  Subtract = number1 - number2
End Function

Function Multiply(number1, number2)
  Multiply = number1 * number2
End Function

Function Divide(number1, number2)
  Divide = number1 / number2
End Function

Function Mod_(number1, number2)
  Mod_ = number1 Mod number2
End Function

Function Power(number, exponent)
  Power = number ^ exponent
End Function

Function Concat(string1, string2)
  Concat = string1 & string2
End Function

Function Not_(expression)
  Not_ = Not expression
End Function

Function And_(expression1, expression2)
  And_ = expression1 And expression2
End Function

Function Or_(expression1, expression2)
  Or_ = expression1 Or expression2
End Function

Function Xor_(expression1, expression2)
  Xor_ = expression1 Xor expression2
End Function


'===================================================
'################ string formatting ################
'---------------------------------------------------

Function LPad(baseString, size, padding)
  Dim result
  result = baseString
  Do While Len(result) < size
    result = padding & result
  Loop
  LPad = result
End Function

Function RPad(baseString, size, padding)
  Dim result
  result = baseString
  Do While Len(result) < size
    result = result & padding
  Loop
  RPad = result
End Function

Dim strftime_TokenScan
Set strftime_TokenScan = re("(.*?)(%.)", "g")

Function strftime(formatExpression, datetime)
  Dim result, lastPos
  Set result = New ListBuffer
  lastPos = 0

  Dim tokenMatch, tokenPrefix, tokenWord
  For Each tokenMatch In strftime_TokenScan.Execute(formatExpression)
    tokenPrefix = tokenMatch.SubMatches(0)
    tokenWord = tokenMatch.SubMatches(1)
    lastPos = tokenMatch.FirstIndex + tokenMatch.Length

    result.Add tokenPrefix
    Select Case tokenWord
      Case "%Y":
        result.Add LPad(DatePart("yyyy", datetime), 4, "0")
      Case "%y":
        result.Add LPad(DatePart("yyyy", datetime) Mod 100, 2, "0")
      Case "%m":
        result.Add LPad(DatePart("m", datetime), 2, "0")
      Case "%d":
        result.Add LPad(DatePart("d", datetime), 2, "0")
      Case "%H":
        result.Add LPad(DatePart("h", datetime), 2, "0")
      Case "%M":
        result.Add LPad(DatePart("n", datetime), 2, "0")
      Case "%S":
        result.Add LPad(DatePart("s", datetime), 2, "0")
      Case "%c":
        result.Add FormatDateTime(datetime)
      Case "%%":
        result.Add "%"
      Case Else:
        result.Add tokenWord
    End Select
  Next

  If lastPos < Len(formatExpression) Then
    result.Add Mid(formatExpression, lastPos + 1)
  End If

  strftime = Join(result.Items, "")
End Function


'======================================
'################ sort ################
'--------------------------------------

Sub SwapArrayItem(list, i, j)
  Dim t
  If IsObject(list(i)) Then
    Set t = list(i)
    Set list(i) = list(j)
    Set list(j) = t
  Else
    t = list(i)
    list(i) = list(j)
    list(j) = t
  End If
End Sub

Sub DownHeap(list, startIndex, maxIndex, compare)
  Dim i, j, k, nextIndex

  i = startIndex
  Do While i <= maxIndex
    j = (i + 1) * 2 - 1
    k = (i + 1) * 2

    If k <= maxIndex Then
      If compare(list(j), list(k)) > 0 Then
        nextIndex = j
      Else
        nextIndex = k
      End If
    ElseIf j <= maxIndex Then
      nextIndex = j
    Else
      Exit Do
    End If

    If compare(list(nextIndex), list(i)) > 0 Then
      SwapArrayItem list, nextIndex, i
    Else
      Exit Do
    End If

    i = nextIndex
  Loop
End Sub

Sub HeapSort(list, compare)
  Dim i

  For i = Int((UBound(list) - 1) / 2) To 0 Step -1
    DownHeap list, i, UBound(list), compare
  Next

  For i = UBound(list) To 1 Step -1
    SwapArrayItem list, 0, i
    DownHeap list, 0, i - 1, compare
  Next
End Sub

Dim Sort
Set Sort = GetRef("HeapSort")


'========================================================
'################ command line arguments ################
'--------------------------------------------------------

Function GetNamedArgumentString(name, namedArgs, default)
  If namedArgs.Exists(name) Then
    If IsEmpty(namedArgs(name)) Then
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentString", "need for value of string option: " & name
    ElseIf VarType(namedArgs(name)) = vbString Then
      GetNamedArgumentString = namedArgs(name)
    Else
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentString", _
         "not a string type named argument: " & name & ":" & ShowValue(namedArgs(name))
    End If
  Else
    If IsEmpty(default) Then
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentString", "need for string option: " & name
    End If
    GetNamedArgumentString = default
  End If
End Function

Function GetNamedArgumentBool(name, namedArgs, default)
  If namedArgs.Exists(name) Then
    If IsEmpty(namedArgs(name)) Then
      GetNamedArgumentBool = True
    ElseIf VarType(namedArgs(name)) = vbBoolean Then
      GetNamedArgumentBool = namedArgs(name)
    Else
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentBool", _
         "not a boolean type named argument: " & name & ":" & ShowValue(namedArgs(name))
    End If
  Else
    If IsEmpty(default) Then
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentBool", "need for boolean option: " & name
    End If
    GetNamedArgumentBool = default
  End If
End Function

Function GetNamedArgumentSimple(name, namedArgs)
  If namedArgs.Exists(name) Then
    If IsEmpty(namedArgs(name)) Then
      GetNamedArgumentSimple = True
    Else
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentSimple", _
         "no need for value of simple option: " & name & ":" & ShowValue(namedArgs(name))
    End If
  Else
    GetNamedArgumentSimple = False
  End If
End Function


'===========================================
'################ file tool ################
'-------------------------------------------

Function FindFile_CreateVisitor
  Dim visitor
  Set visitor = D(Array("fso", CreateObject("Scripting.FileSystemObject")))

  PseudoObject_AttachMethodSubProc visitor, "TraverseDrive", GetRef("FindFile_TraverseDrive"), 2
  PseudoObject_AttachMethodSubProc visitor, "TraverseFolder", GetRef("FindFile_TraverseFolder"), 2

  PseudoObject_AttachMethodSubProc _
              visitor, "TraverseDrive_ErrorHandler", GetRef("FindFile_TraverseDrive_ErrorHandlerDefault"), 3
  PseudoObject_AttachMethodSubProc _
              visitor, "TraverseFolder_ErrorHandler", GetRef("FindFile_TraverseFolder_ErrorHandlerDefault"), 3

  PseudoObject_AttachMethodSubProc visitor, "VisitDrive", GetRef("FindFile_VisitDriveDefault"), 2
  PseudoObject_AttachMethodSubProc visitor, "VisitFolder", GetRef("FindFile_VisitFolderDefault"), 2
  PseudoObject_AttachMethodSubProc visitor, "VisitFile", GetRef("FindFile_VisitFileDefault"), 2

  Set FindFile_CreateVisitor = visitor
End Function

Sub FindFile_PathAccept(path, visitor)
  Dim fso
  Set fso = visitor("fso")
  If fso.DriveExists(path) Then
    visitor("VisitDrive")(fso.GetDrive(path))
  ElseIf fso.FolderExists(path) Then
    visitor("VisitFolder")(fso.GetFolder(path))
  ElseIf fso.FileExists(path) Then
    visitor("VisitFile")(fso.GetFile(path))
  Else
    Err.Raise RuntimeError, "not exists: " & ShowValue(path)
  End If
End Sub

Sub FindFile_AllDriveAccept(visitor)
  Dim fso, drive
  Set fso = visitor("fso")
  For Each drive In fso.Drives
    visitor("VisitDrive")(drive)
  Next
End Sub

Sub FindFile_TraverseDrive(self, drive)
  Dim isReady, rootFolder, errorContext
  On Error Resume Next

  isReady = drive.IsReady
  If Err.Number <> 0 Then
    Set errorContext = D(Array("Number", Err.Number, _
                               "Source", Err.Source, _
                               "Description", Err.Description))
    Err.Clear
    On Error GoTo 0
    Call (self("TraverseDrive_ErrorHandler"))(drive, errorContext)
    On Error Resume Next
    Exit Sub
  End If

  Set rootFolder = drive.RootFolder
  If Err.Number <> 0 Then
    Set errorContext = D(Array("Number", Err.Number, _
                               "Source", Err.Source, _
                               "Description", Err.Description))
    Err.Clear
    On Error GoTo 0
    Call (self("TraverseDrive_ErrorHandler"))(drive, errorContext)
    On Error Resume Next
    Exit Sub
  End If

  On Error GoTo 0

  If isReady Then
    self("VisitFolder")(rootFolder)
  End If
End Sub

Sub FindFile_TraverseFolder(self, folder)
  Dim f
  Dim errorContext
  Set errorContext = Nothing

  Err.Clear
  On Error Resume Next

  For Each f In folder.Files
    If Err.Number <> 0 Then
      Set errorContext = D(Array("Number", Err.Number, _
                                 "Source", Err.Source, _
                                 "Description", Err.Description))
      Err.Clear
      On Error GoTo 0
      Call (self("TraverseFolder_ErrorHandler"))(folder, errorContext)
      On Error Resume Next
      Exit For
    End If

    On Error GoTo 0
    self("VisitFile")(f)
    On Error Resume Next
  Next

  For Each f In folder.SubFolders
    If Err.Number <> 0 Then
      If Not errorContext Is Nothing Then
        If Err.Number = errorContext("Number") And _
           Err.Source = errorContext("Source") And _
           Err.Description = errorContext("Description") _
        Then
          Exit For                      ' skip same error
        End If
      End If

      Set errorContext = D(Array("Number", Err.Number, _
                                 "Source", Err.Source, _
                                 "Description", Err.Description))
      Err.Clear
      On Error GoTo 0
      Call (self("TraverseFolder_ErrorHandler"))(folder, errorContext)
      On Error Resume Next
      Exit For
    End If

    On Error GoTo 0
    self("VisitFolder")(f)
    On Error Resume Next
  Next
End Sub

Sub FindFile_TraverseDrive_ErrorHandlerDefault(self, drive, errorContext)
  Err.Raise RuntimeError, "stdlib.vbs:TraverseDrive_ErrorHandlerDefault", _
     "failed to access drive: " & drive.Path & vbNewLine & _
     "<" & errorContext("Number") & "> " & errorContext("Description") & " (" & errorContext("Source") & ")"
End Sub

Sub FindFile_TraverseFolder_ErrorHandlerDefault(self, folder, errorContext)
  Err.Raise RuntimeError, "stdlib.vbs:TraverseFolder_ErrorHandlerDefault", _
     "failed to access folder: " & folder.Path & vbNewLine & _
     "<" & errorContext("Number") & "> " & errorContext("Description") & " (" & errorContext("Source") & ")"
End Sub

Sub FindFile_VisitDriveDefault(self, drive)
  self("TraverseDrive")(drive)
End Sub

Sub FindFile_VisitFolderDefault(self, folder)
  self("TraverseFolder")(folder)
End Sub

Sub FindFile_VisitFileDefault(self, file)
End Sub

Dim ZipFile_EmptyData
ZipFile_EmptyData = "PK" & Chr(&H05) & Chr(&H06) & String(18, Chr(&H00))

Dim ZipFile_Extension
Set ZipFile_Extension = re("\.zip$", "i")

Class ZipFileObject
  Private ivar_fso
  Private ivar_shellApp
  Private ivar_timeoutSeconds
  Private ivar_pollingIntervalMillisecs

  Private Sub Class_Initialize
    Set ivar_fso = CreateObject("Scripting.FileSystemObject")
    Set ivar_shellApp = CreateObject("Shell.Application")
    ivar_timeoutSeconds = 60
    ivar_pollingIntervalMillisecs = 100
  End Sub

  Public Property Get TimeoutSeconds
    TimeoutSeconds = ivar_timeoutSeconds
  End Property

  Public Property Let TimeoutSeconds(seconds)
    ivar_timeoutSeconds = seconds
  End Property

  Public Property Get PollingIntervalMillisecs
    PollingIntervalMillisecs = ivar_pollingIntervalMillisecs
  End Property

  Public Property Let PollingIntervalMillisecs(milliseconds)
    ivar_pollingIntervalMillisecs = milliseconds
  End Property

  Public Function IsOpened(filename)
    Dim f
    Dim errNum, errSrc, errDsc

    On Error Resume Next
    Set f = ivar_fso.OpenTextFile(filename, 8, False) ' 8 for Appending
    errNum = Err.Number
    errSrc = Err.Source
    errDsc = Err.Description
    Err.Clear
    On Error GoTo 0

    Select Case errNum
      Case 0:
        f.Close
        IsOpened = False
      Case 70:
        IsOpened = True
      Case Else:
        Err.Raise errNum, errSrc, errDsc
    End Select
  End Function

  Public Sub CreateEmptyZipFile(filename, overwrite)
    With ivar_fso.CreateTextFile(filename, overwrite)
      .Write ZipFile_EmptyData
      .Close
    End With
  End Sub

  Private Function GetZipName(filename)
    Dim zipInfo
    Set zipInfo = D(Array("Name", filename))
    If Not ZipFile_Extension.Test(zipInfo("Name")) Then
      zipInfo("Name") = zipInfo("Name") & ".zip"
    End If
    zipInfo("AbsPath") = ivar_fso.GetAbsolutePathName(zipInfo("Name"))
    Set GetZipName = zipInfo
  End Function

  ' Success -> True
  ' Failure -> False
  Private Function WaitForItemsChanged(zipPath, itemsCount)
    Dim startTime
    startTime = Now

    Dim zipFolder
    Do
      Set zipFolder = ivar_shellApp.NameSpace(zipPath) ' need to create new zip folder object in updating.
      If Not zipFolder Is Nothing Then
        If zipFolder.Items.Count <> itemsCount Then
          Exit Do
        End If
      End If
      Set zipFolder = Nothing           ' need to release zip folder object in updating.

      If DateDiff("s", startTime, Now) > ivar_timeoutSeconds Then
        WaitForItemsChanged = False
        Exit Function
      End If
      WScript.Sleep ivar_pollingIntervalMillisecs
    Loop
    Set zipFolder = Nothing             ' need to release zip folder object in updating.

    Do While IsOpened(zipPath)
      If DateDiff("s", startTime, Now) > ivar_timeoutSeconds Then
        WaitForItemsChanged = False
        Exit Function
      End If
      WScript.Sleep ivar_pollingIntervalMillisecs
    Loop

    WaitForItemsChanged = True
  End Function

  Public Sub Zip(filename, entries)
    Dim z
    Set z = GetZipName(filename)

    CreateEmptyZipFile z("Name"), True

    If ivar_shellApp.NameSpace(z("AbsPath")) Is Nothing Then
      Err.Raise RuntimeError, "stdlib.vbs:ZipFileObject.Zip", _
         "not found a zip file: " & z("Name")
    End If

    Dim entryName
    For Each entryName In entries
      If Not ivar_fso.FileExists(entryName) And Not ivar_fso.FolderExists(entryName) Then
        Err.Raise RuntimeError, "stdlib.vbs:ZipFileObject.Zip", _
           "not found file or folder: " & entryName
      End If

      Dim zipFolder, count
      Set zipFolder = ivar_shellApp.NameSpace(z("AbsPath")) ' need to create new zip folder object in each update.
      count = zipFolder.Items.Count
      zipFolder.CopyHere ivar_fso.GetAbsolutePathName(entryName)
      Set zipFolder = Nothing           ' need to release zip folder object in each update.

      If Not WaitForItemsChanged(z("AbsPath"), count) Then
        Err.Raise RuntimeError, "stdlib.vbs:ZipFileObject.Zip", _
           "failed to add an entry to zip file: " & entryName & " -> " & z("Name")
      End If
    Next
  End Sub

  Public Sub Unzip(filename, destFolderPath)
    Dim z
    Set z = GetZipName(filename)

    Dim zipFolder
    Set zipFolder = ivar_shellApp.NameSpace(z("AbsPath"))
    If zipFolder Is Nothing Then
      Err.Raise RuntimeError, "stdlib.vbs:ZipFileObject.Unzip", _
         "not found a zip file: " & z("Name")
    End If

    Dim destFolder
    Set destFolder = ivar_shellApp.NameSpace(ivar_fso.GetAbsolutePathName(destFolderPath))
    If destFolder Is Nothing Then
      Err.Raise RuntimeError, "stdlib.vbs:ZipFileObject.Unzip", _
         "not found a folder: " & destFolderPath
    End If

    Dim item
    For Each item In zipFolder.Items
      destFolder.CopyHere item
    Next
  End Sub

  Private Function ZipFolderEntries(parentPath, zipFolder)
    Dim entryList
    Set entryList = New ListBuffer

    Dim item, itemPath
    For Each item In zipFolder.Items
      itemPath = ivar_fso.BuildPath(parentPath, item.Name)
      If item.IsFolder Then
        entryList.Append ZipFolderEntries(itemPath, item.GetFolder)
      Else
        entryList.Add itemPath
      End If
    Next

    ZipFolderEntries = entryList.Items
  End Function

  Public Function ZipEntries(filename)
    Dim z
    Set z = GetZipName(filename)

    Dim zipFolder
    Set zipFolder = ivar_shellApp.NameSpace(z("AbsPath"))
    If zipFolder Is Nothing Then
      Err.Raise RuntimeError, "stdlib.vbs:ZipFileObject.ZipEntries", _
         "not found a zip file: " & z("Name")
    End If

    ZipEntries = ZipFolderEntries(z("Name"), zipFolder)
  End Function
End Class


'==========================================
'################ GUI tool ################
'------------------------------------------

Class FileOpenDialog
  Private ivar_ie

  Private Sub Class_Initialize
    Set ivar_ie = CreateObject("InternetExplorer.Application")
    ivar_ie.MenuBar = False
    ivar_ie.AddressBar = False
    ivar_ie.ToolBar = False
    ivar_ie.StatusBar = False
    ivar_ie.Navigate "about:blank"
    'ivar_ie.Visible = True
    WaitReadyStateComplete
    ivar_ie.document.Write "<html><body></body></html>"
  End Sub

  Private Sub Class_Terminate
    ivar_ie.Quit
    Set ivar_ie = Nothing
  End Sub

  Private Sub WaitReadyStateComplete
    Do While ivar_ie.Busy And ivar_ie.ReadyState <> 4
      WScript.Sleep 10
    Loop
  End Sub

  Public Function GetFilePath
    ivar_ie.document.body.innerHTML = "<input type='file' id='FileOpenDialog' />"
    Dim file
    Set file = ivar_ie.document.getElementById("FileOpenDialog")
    file.Click
    If Len(file.Value) > 0 Then
      GetFilePath = file.Value
    End If
  End Function
End Class

Function InputFileOpenDialog
  Dim dialog
  Set dialog = New FileOpenDialog
  InputFileOpenDialog = dialog.GetFilePath
End Function

Function InputFolderDialog(title)
  Dim s, folder
  Set s = CreateObject("Shell.Application")
  Set folder = s.BrowseForFolder(0, title, &H0040)
  If Not folder Is Nothing Then
    InputFolderDialog = folder.Self.Path
  End If
End Function


'===========================================
'################ ADSI tool ################
'-------------------------------------------

Function ADSI_CreateVisitor
  Dim visitor
  Set visitor = D(Array("__schemaCache__", CreateObject("Scripting.Dictionary")))
  PseudoObject_AttachMethodFuncProc visitor, "GetSchema", GetRef("ADSI_GetSchema"), 2
  PseudoObject_AttachMethodFuncProc visitor, "IsContainer", GetRef("ADSI_IsContainer"), 2
  PseudoObject_AttachMethodSubProc visitor, "ADSI_VisitObject", GetRef("ADSI_VisitObjectDefault"), 2
  PseudoObject_AttachMethodSubProc visitor, "ADSI_VisitContainer", GetRef("ADSI_VisitContainerDefault"), 2
  Set ADSI_CreateVisitor = visitor
End Function

Sub ADSI_Accept(adsObject, visitor)
  Dim key
  key = "ADSI_Visit_" & adsObject.Class
  If visitor.Exists(key) Then
    visitor(key)(adsObject)
  Else
    If visitor("IsContainer")(adsObject) Then
      visitor("ADSI_VisitContainer")(adsObject)
    Else
      visitor("ADSI_VisitObject")(adsObject)
    End If
  End If
End Sub

Function ADSI_GetSchema(self, adsObject)
  Dim schemaCache
  Set schemaCache = self("__schemaCache__")
  If Not schemaCache.Exists(adsObject.Schema) Then
    schemaCache.Add adsObject.Schema, GetObject(adsObject.Schema)
  End If
  Set ADSI_GetSchema = schemaCache(adsObject.Schema)
End Function

Function ADSI_IsSchema(adsObject)
  Select Case adsObject.Class
    Case "Schema":
      ADSI_IsSchema = True
    Case "Class":
      ADSI_IsSchema = True
    Case "Syntax":
      ADSI_IsSchema = True
    Case "Property":
      ADSI_IsSchema = True
    Case Else:
      ADSI_IsSchema = False
  End Select
End Function

Function ADSI_IsContainer(self, adsObject)
  If ADSI_IsSchema(adsObject) Then
    ADSI_IsContainer = False
  Else
    ADSI_IsContainer = self("GetSchema")(adsObject).Container
  End If
End Function

Sub ADSI_TraverseCollection(self, adsCollection)
  Dim adsObject
  For Each adsObject In adsCollection
    ADSI_Accept adsObject, self
  Next
End Sub

Sub ADSI_VisitObjectDefault(self, adsObject)
End Sub

Sub ADSI_VisitContainerDefault(self, adsContainer)
  ADSI_TraverseCollection self, adsContainer
End Sub


'==========================================
'################ WMI tool ################
'------------------------------------------

Class WbemPropertyOptionalInformationGetter
  Private ivar_class
  Private ivar_cache

  Private Sub Class_Initialize
    Set ivar_cache = CreateObject("Scripting.Dictionary")
  End Sub

  Public Property Set [Class](value)
    Set ivar_class = value
  End Property

  Private Function GetQualifier(prop, Name)
    Err.Clear
    On Error Resume Next
    Set GetQualifier = prop.Qualifiers_(name)
    If Err.Number <> 0 Then
      Set GetQualifier = Nothing
      Err.Clear
    End If
  End Function

  Public Default Function Execute(propName)
    If Not ivar_cache.Exists(propName) Then
      Dim prop
      Set prop = ivar_Class.Properties_(propName)

      Dim units
      Set units = GetQualifier(prop, "Units")

      Dim valueMap
      Set valueMap = GetQualifier(prop, "ValueMap")

      Dim info: info = ""
      If Not units Is Nothing Then
        info = info & " (" & units.Value & ")"
      End If
      If Not valueMap Is Nothing Then
        info = info & " [" & Join(valueMap.Value, "|") & "]"
      End If

      ivar_cache.Add propName, info
    End If

    Execute = ivar_cache(propName)
  End Function
End Class


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
