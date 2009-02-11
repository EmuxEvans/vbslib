' VBScript Standard Library

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

Sub BindAt(toStore, index, value)
  If IsObject(value) Then
    Set toStore(index) = value
  Else
    toStore(index) = value
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

Function re(regexpPattern, regexpOptions)
  Dim regex
  Set regex = New RegExp
  regex.Pattern = regexpPattern
  regexpOptions = LCase(regexpOptions)
  If InStr(regexpOptions, "i") > 0 Then
    regex.IgnoreCase = True
  End If
  If InStr(regexpOptions, "g") > 0 Then
    regex.Global = True
  End If
  If InStr(regexpOptions, "m") > 0 Then
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
    BindAt ivar_dict, nextIndex, value
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
End Class

Dim ShowString_Quote
Set ShowString_Quote = re("""", "g")

Function ShowString(value)
  ShowString = """" & ShowString_Quote.Replace(value, """""") & """"
End Function

Function ShowArray(value)
  Dim r, i, sep: sep = ""
  r = "["
  For Each i In value
    r = r & sep & ShowValue(i)
    sep = ","
  Next
  r = r & "]"
  ShowArray = r
End Function

Function ShowDictionary(value)
  Dim r, k, sep: sep = ""
  r = "{"
  For Each k In value.Keys
    r = r & sep & ShowValue(k) & "=>" & ShowValue(value(k))
    sep = ","
  Next
  r = r & "}"
  ShowDictionary = r
End Function

Function ShowObject(value)
  On Error Resume Next
  Dim r
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
  On Error Resume Next
  Dim r
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
    r = "<unknown:" & VarType(value) & ">"
  End If
  ShowOther = r
End Function

Function ShowValue(value)
  Dim r
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
  ShowValue = r
End Function


'===========================================
'################ list tool ################
'-------------------------------------------

Function Find(list, cond)
  Dim i
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

Function Equal(expected, value)
  If value = expected Then
    Equal = True
  Else
    Equal = False
  End If
End Function

Function ValueEqual(expected)
  Set ValueEqual = _
      GetFuncProcSubset(GetRef("Equal"), 2, Array(expected))
End Function

Function GreaterThan(lowerBound, value)
  If value > lowerBound Then
    GreaterThan = True
  Else
    GreaterThan = False
  End If
End Function

Function GreaterThanEqual(lowerBound, value)
  If value >= lowerBound Then
    GreaterThanEqual = True
  Else
    GreaterThanEqual = False
  End If
End Function

Function ValueGreaterThan(lowerBound, exclude)
  If exclude Then
    Set ValueGreaterThan = _
        GetFuncProcSubset(GetRef("GreaterThan"), 2, Array(lowerBound))
  Else
    Set ValueGreaterThan = _
        GetFuncProcSubset(GetRef("GreaterThanEqual"), 2, Array(lowerBound))
  End If
End Function

Function LessThan(upperBound, value)
  If value < upperBound Then
    LessThan = True
  Else
    LessThan = False
  End If
End Function

Function LessThanEqual(upperBound, value)
  If value <= upperBound Then
    LessThanEqual = True
  Else
    LessThanEqual = False
  End If
End Function

Function ValueLessThan(upperBound, exclude)
  If exclude Then
    Set ValueLessThan = _
        GetFuncProcSubset(GetRef("LessThan"), 2, Array(upperBound))
  Else
    Set ValueLessThan = _
        GetFuncProcSubset(GetRef("LessThanEqual"), 2, Array(upperBound))
  End If
End Function

Function Between(lowerBound, upperBound, value)
  If (lowerBound <= value) And (value <= upperBound) Then
    Between = True
  Else
    Between = False
  End If
End Function

Function BetweenExcludeUpperBound(lowerBound, upperBound, value)
  If (lowerBound <= value) And (value < upperBound) Then
    BetweenExcludeUpperBound = True
  Else
    BetweenExcludeUpperBound = False
  End If
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

Function NotFunc(cond, value)
  If Not cond(value) Then
    NotFunc = True
  Else
    NotFunc = False
  End If
End Function

Function NotCond(cond)
  Set NotCond = GetFuncProcSubset(GetRef("NotFunc"), 2, Array(cond))
End Function

Function AndFunc(cond1, cond2, value)
  If cond1(value) And cond2(value) Then
    AndFunc = True
  Else
    AndFunc = False
  End If
End Function

Function AndCond(cond1, cond2)
  Set AndCond = GetFuncProcSubset(GetRef("AndFunc"), 3, Array(cond1, cond2))
End Function

Function OrFunc(cond1, cond2, value)
  If cond1(value) Or cond2(value) Then
    OrFunc = True
  Else
    OrFunc = False
  End If
End Function

Function OrCond(cond1, cond2)
  Set OrCond = GetFuncProcSubset(GetRef("OrFunc"), 3, Array(cond1, cond2))
End Function

Function Max(list, compare)
  Dim first
  first = True

  Dim x, maxValue
  For Each x In list
    If first Then
      Bind maxValue, x
    Else
      If compare(x, maxValue) > 0 Then
        Bind maxValue, x
      End If
    End If
    first = False
  Next

  Bind Max, maxValue
End Function

Function Min(list, compare)
  Dim first
  first = True

  Dim x, minValue
  For Each x In list
    If first Then
      Bind minValue, x
    Else
      If compare(x, minValue) < 0 Then
        Bind minValue, x
      End If
    End If
    first = False
  Next

  Bind Min, minValue
End Function

Function NumberCompare(a, b)
  NumberCompare = a - b
End Function

Function TextStringCompare(a, b)
  TextStringCompare = StrComp(a, b, vbTextCompare)
End Function

Function BinaryStringCompare(a, b)
  BinaryStringCompare = StrComp(a, b, vbBinaryCompare)
End Function

Function ObjectPropertyCompare_(propName, propComp, a, b)
  ObjectPropertyCompare_ = propComp(GetObjectProperty(a, propName), _
                                    GetObjectProperty(b, propName))
End Function

Function ObjectPropertyCompare(propertyName, propertyCompare)
  Set ObjectPropertyCompare = _
      GetFuncProcSubset(GetRef("ObjectPropertyCompare_"), 4, Array(propertyName, propertyCompare))
End Function

Function Map(list, func)
  Dim newList, i
  Set newList = New ListBuffer
  For Each i In list
    newList.Add func(i)
  Next
  Map = newList.Items
End Function

Function ValueReplace(regex, replace)
  Set ValueReplace = _
      GetFuncProcSubset(GetObjectMethodFuncProc(regex, "Replace", 2), 2, D(Array(1, replace)))
End Function

Function ValueObjectProperty(propertyName)
  Set ValueObjectProperty = _
      GetFuncProcSubset(GetRef("GetObjectProperty"), 2, D(Array(1, propertyName)))
End Function

Function GetDictionaryItem(key, dictionary)
  If dictionary.Exists(key) Then
    Bind GetDictionaryItem, dictionary(key)
  End If
End Function

Function ValueDictionaryItem(key)
  Set ValueDictionaryItem = _
      GetFuncProcSubset(GetRef("GetDictionaryItem"), 2, Array(key))
End Function

' aliases of the VBScript functions to GetRef().
Sub MapFunc_DefineAliases
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
  aliases.Add Array("TypeName", 1)
  aliases.Add Array("VarType", 1)

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

  ' Number
  aliases.Add Array("Hex", 1)
  aliases.Add Array("Oct", 1)
  aliases.Add Array("Sgn", 1)
  aliases.Add Array("Int", 1)
  aliases.Add Array("Fix", 1)

  ' DateTime
  aliases.Add Array("DateValue", 1)
  aliases.Add Array("TimeValue", 1)
  aliases.Add Array("Year", 1)
  aliases.Add Array("Day", 1)
  aliases.Add Array("Month", 1)
  aliases.Add Array("Hour", 1)
  aliases.Add Array("Minute", 1)
  aliases.Add Array("Second", 1)

  ' Eval
  aliases.Add Array("Eval", 1)
  aliases.Add Array("GetRef", 1)

  Dim aliasPair, aliasExpr, name, argCount, argList, sep, i
  For Each aliasPair In aliases.Items
    name = aliasPair(0)
    argCount = aliasPair(1)

    argList = ""
    sep = ""
    For i = 0 To argCount - 1
      argList = argList & sep & "arg" & i
      sep = ", "
    Next

    Set aliasExpr = New ListBuffer
    aliasExpr.Add "Function " & name & "_(" & argList & ")"
    aliasExpr.Add "  Bind " & name & "_, " & name & "(" & argList & ")"
    aliasExpr.Add "End Function"

    ExecuteGlobal Join(aliasExpr.Items, vbNewLine)
  Next
End Sub

MapFunc_DefineAliases


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
    Set ObjectProperty_AccessorPool(key) = ObjectProperty_CreateAccessor(name)
  End If
  Set ObjectProperty_GetAccessor = ObjectProperty_AccessorPool(key)
End Function

Function GetObjectProperty(obj, name)
  Bind GetObjectProperty, ObjectProperty_GetAccessor(name)(obj)
End Function

Sub SetObjectProperty(obj, name, value)
  BindAt ObjectProperty_GetAccessor(name), obj, value
End Sub

Function ExistsObjectProperty(obj, name)
  On Error Resume Next
  ObjectProperty_GetAccessor(name)(obj)
  Select Case Err.Number
    Case 0:
      Err.Clear
      ExistsObjectProperty = True
    Case 438:
      Err.Clear
      ExistsObjectProperty = False
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
  classExpr.Add "  Public Sub InvokeMethod(obj, args)"
  classExpr.Add "    obj." & name & " " & argList
  classExpr.Add "  End Sub"
  classExpr.Add ""
  classExpr.Add "  Public Function FuncallMethod(obj, args)"
  classExpr.Add "    Bind FuncallMethod, obj." & name & "(" & argList & ")"
  classExpr.Add "  End Function"
  classExpr.Add "End Class"

  ExecuteGlobal Join(classExpr.Items, vbNewLine)
  Set ObjectMethod_CreateHandler = Eval("New " & className)
End Function

Function ObjectMethod_GetHandler(name, argCount)
  Dim key: key = UCase(name) & "_" & argCount
  If Not ObjectMethod_HandlerPool.Exists(key) Then
    Set ObjectMethod_HandlerPool(key) = ObjectMethod_CreateHandler(name, argCount)
  End If
  Set ObjectMethod_GetHandler = ObjectMethod_HandlerPool(key)
End Function

Sub InvokeObjectMethod(obj, name, args)
  Dim argCount, handler
  If IsArray(args) Then
    argCount = UBound(args) + 1
  ElseIf IsObject(args) Then
    argCount = args.Count
  Else
    Err.Raise 13, "stdlib.vbs:InvokeObjectMethod", "args is not Array."
  End If
  Set handler = ObjectMethod_GetHandler(name, argCount)
  handler.InvokeMethod obj, args
End Sub

Function FuncallObjectMethod(obj, name, args)
  Dim argCount, handler
  If IsArray(args) Then
    argCount = UBound(args) + 1
  ElseIf IsObject(args) Then
    argCount = args.Count
  Else
    Err.Raise 13, "stdlib.vbs:FuncallObjectMethod", "args is not Array."
  End If
  Set handler = ObjectMethod_GetHandler(name, argCount)
  Bind FuncallObjectMethod, handler.FuncallMethod(obj, args)
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
  classExpr.Add "  Public Property Set self(obj)"
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
    Set ObjectMethod_ProcBuilderPool(key) = ObjectMethod_CreateProcBuilder(name, argCount)
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
  key = "arg" & argCount & "_" & Join(paramIndexList, " ")
  If Not ProcSubset_ProcBuilderPool.Exists(key) Then
    Set ProcSubset_ProcBuilderPool(key) = ProcSubset_CreateProcBuilder(argCount, paramIndexList)
  End If
  Set ProcSubset_GetProcBuilder = ProcSubset_ProcBuilderPool(key)
End Function

Dim ProcSubset_NumberCompare
Set ProcSubset_NumberCompare = GetRef("NumberCompare")

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
  Sort paramIndexList, ProcSubset_NumberCompare

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
  If namedArgs.Exsits(name) Then
    If IsEmpty(namedArgs(name)) Then
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentString", "need for value of string option: " & name
    ElseIf VarType(namedArgs(name)) = vbString Then
      GetNamedArgumentString = namedArgs(name)
    Else
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentString", _
        "not a string type named argument: " & name & ": " & ShowValue(namedArgs(name))
    End If
  Else
    If default Is Nothing Then
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
    If default Is Nothing Then
      Err.Raise RuntimeError, "stdlib.vbs:GetNamedArgumentBool", "need for boolean option: " & name
    End If
    GetNamedArgumentBool = default
  End If
End Function


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
