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
  className = "ObjectMethod_Handler_" & name & "_" & argCount
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
  className = "ObjectMethod_Proc_" & name & "_" & argCount
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


'=============================================
'################ Display I/O ################
'---------------------------------------------

Class ConsoleWriter
  Public Sub Write(message)
    WScript.StdOut.Write(message)
  End Sub

  Public Sub Flush
  End Sub

  Public Sub FlushAndWrite(lastMessage)
    Write(lastMessage)
  End Sub
End Class

Class MsgBoxWriter
  Private ivar_buffer

  Private Sub Class_Initialize
    Set ivar_buffer = New ListBuffer
  End Sub

  Public Sub Write(message)
    ivar_buffer.Add message
  End Sub

  Public Sub Flush
    Dim s: s = ""
    Dim msg
    For Each msg In ivar_buffer.Items
      s = s & msg
    Next
    ivar_buffer.Clear
    MsgBox s, vbOKOnly + vbInformation, WScript.ScriptName
  End Sub

  Public Sub FlushAndWrite(lastMessage)
    Flush
  End Sub
End Class

Class MessageWriter
  Private out

  Private Sub Class_Initialize
    Set out = New ConsoleWriter
  End Sub

  Public Sub Write(message)
    On Error Resume Next
    out.Write message
    If Err.Number <> 0 Then
      Err.Clear
      On Error GoTo 0
      out = New MsgBoxWriter
      out.Write message
    End If
  End Sub

  Public Default Sub WriteLine(message)
    Write message & vbNewLine
  End Sub

  Public Sub Flush
    out.Flush
  End Sub

  Public Sub FlushAndWrite(lastMessage)
    On Error Resume Next
    out.FlushAndWrite lastMessage
    If Err.Number <> 0 Then
      Err.Clear
      On Error GoTo 0
      out = New MsgBoxWriter
      out.FlushAndWrite lastMessage
    End If
  End Sub

  Public Sub FlushAndWriteLine(lastMessage)
    FlushAndWrite lastMessage & vbNewLine
  End Sub
End Class

Dim MsgOut
Set MsgOut = New MessageWriter


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
