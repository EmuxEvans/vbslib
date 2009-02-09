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

  Dim count, i, key
  count = 0

  For Each i In keyValueList
    If (count Mod 2) = 0 Then
      Bind key, i
      dict.Add key, Empty
    Else
      BindAt dict, key, i
    End If
    count = count + 1
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
Set ShowString_Quote = New RegExp
ShowString_Quote.Pattern = """"
ShowString_Quote.Global = True

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

Class ValueEqualCondition
  Private ivar_expectedValue

  Public Property Let ExpectedValue(value)
    ivar_expectedValue = value
  End Property

  Public Default Function Apply(value)
    If value = ivar_expectedValue Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Function ValueEqual(expectedValue)
  Dim cond
  Set cond = New ValueEqualCondition
  cond.Expectedvalue = expectedvalue
  Set ValueEqual = cond
End Function

Class ValueGreaterThanCondition
  Private ivar_lowerBound

  Public Property Let LowerBound(value)
    ivar_lowerBound = value
  End Property

  Public Default Function Apply(value)
    If value > ivar_lowerBound Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Class ValueGreaterThanEqualCondition
  Private ivar_lowerBound

  Public Property Let LowerBound(value)
    ivar_lowerBound = value
  End Property

  Public Default Function Apply(value)
    If value >= ivar_lowerBound Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Function ValueGreaterThan(lowerBound, exclude)
  Dim cond
  If exclude Then
    Set cond = New ValueGreaterThanCondition
  Else
    Set cond = New ValueGreaterThanEqualCondition
  End If
  cond.Lowerbound = lowerBound
  Set ValueGreaterThan = cond
End Function

Class ValueLessThanCondition
  Private ivar_upperBound

  Public Property Let UpperBound(value)
    ivar_upperBound = value
  End Property

  Public Default Function Apply(value)
    If value < ivar_upperBound Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Class ValueLessThanEqualCondition
  Private ivar_upperBound

  Public Property Let UpperBound(value)
    ivar_upperBound = value
  End Property

  Public Default Function Apply(value)
    If value <= ivar_upperBound Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Function ValueLessThan(upperBound, exclude)
  Dim cond
  If exclude Then
    Set cond = New ValueLessThanCondition
  Else
    Set cond = New ValueLessThanEqualCondition
  End If
  cond.UpperBound = upperBound
  Set ValueLessThan = cond
End Function

Class ValueBetweenCondition
  Private ivar_lowerBound
  Private ivar_upperBound

  Public Property Let LowerBound(value)
    ivar_lowerBound = value
  End Property

  Public Property Let UpperBound(value)
    ivar_upperBound = value
  End Property

  Public Default Function Apply(value)
    If (ivar_lowerBound <= value) And (value <= ivar_upperBound) Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Class ValueBetweenExcludeUpperBoundCondition
  Private ivar_lowerBound
  Private ivar_upperBound

  Public Property Let LowerBound(value)
    ivar_lowerBound = value
  End Property

  Public Property Let UpperBound(value)
    ivar_upperBound = value
  End Property

  Public Default Function Apply(value)
    If (ivar_lowerBound <= value) And (value < ivar_upperBound) Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Function ValueBetween(lowerBound, upperBound, exclude)
  Dim cond
  If exclude Then
    Set cond = New ValueBetweenExcludeUpperBoundCondition
  Else
    Set cond = New ValueBetweenCondition
  End If
  cond.LowerBound = lowerBound
  cond.UpperBound = upperBound
  Set ValueBetween = cond
End Function

Class RegExpMatchCondition
  Private ivar_regexp

  Public Property Set RegExp(value)
    Set ivar_regexp = value
  End Property

  Public Default Function Apply(value)
    If ivar_regexp.Test(value) Then
      Apply = True
    Else
      Apply = False
    End If
  End Function
End Class

Function RegExpMatch(regex)
  Dim cond
  Set cond = New RegExpMatchCondition
  Set cond.RegExp = regex
  Set RegExpMatch = cond
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
  classExpr.Add "  Public Default Sub Apply(" & argList & ")"
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
  classExpr.Add "  Public Default Function Apply(" & argList & ")"
  classExpr.Add "    Bind Apply, ivar_obj." & name & "(" & argList & ")"
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

Function ProcSubset_CreateProcBuilder(argCount, paramIndexList)
  Dim className, classExpr
  className = "ProcSubset_Arg" & argCount & "_" & Join(paramIndexList, "_")
  Set classExpr = New ListBuffer

  Dim argList, applyArgList, i
  Set argList = New ListBuffer
  Set applyArgList = New ListBuffer
  For i = 0 To argCount - 1
    If IsEmpty(Find(paramIndexList, ValueEqual(i))) Then
      argList.Add "arg" & i
      applyArgList.Add "arg" & i
    Else
      applyArgList.Add "ivar_arg" & i
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
  classExpr.Add "  Public Default Sub Apply(" & argList & ")"
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
  classExpr.Add "  Public Default Function Apply(" & argList & ")"
  classExpr.Add "    Bind Apply, ivar_proc(" & applyArgList & ")"
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
  Sort paramIndexList, NumberCompare

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

Dim NumberCompare
Set NumberCompare = GetRef("NumberCompareFunction")

Function NumberCompareFunction(a, b)
  NumberCompareFunction = a - b
End Function

Dim TextStringCompare
Set TextStringCompare = GetRef("TextStringCompareFunction")

Function TextStringCompareFunction(a, b)
  TextStringCompareFunction = StrComp(a, b, vbTextCompare)
End Function

Dim BinaryStringCompare
Set BinaryStringCompare = GetRef("BinaryStringCompareFunction")

Function BinaryStringCompareFunction(a, b)
  BinaryStringCompareFunction = StrComp(a, b, vbBinaryCompare)
End Function

Class ObjectPropertyCompare
  Private ivar_propName
  Private ivar_propComp

  Public Property Let PropertyName(value)
    ivar_propName = value
  End Property

  Public Property Set PropertyCompare(value)
    Set ivar_propComp = value
  End Property

  Public Default Function Compare(a, b)
    If IsEmpty(ivar_propName) Then
      Err.Raise RuntimeError, "stdlib.vbs:ObjectPropertyCompare", "Not defined `PropertyName'."
    End If
    If IsEmpty(ivar_propComp) Then
      Err.Raise RuntimeError, "stdlib.vbs:ObjectPropertyCompare", "Not defined `PropertyCompare'."
    End If
    Compare = ivar_propComp(GetObjectProperty(a, ivar_propName), _
                            GetObjectProperty(b, ivar_propName))
  End Function
End Class

Function CreateObjectPropertyCompare(propertyName, propertyCompare)
  Dim compare
  Set compare = New ObjectPropertyCompare
  compare.PropertyName = propertyName
  Set compare.PropertyCompare = propertyCompare
  Set New_ObjectPropertyCompare = compare
End Function

'========================================================
'################ command line arguments ################
'--------------------------------------------------------

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
    GetNamedArgumentBool = default
  End If
End Function


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
