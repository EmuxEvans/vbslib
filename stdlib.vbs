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

  Public Sub Add(value)
    Dim nextIndex
    nextIndex = ivar_dict.Count
    BindAt ivar_dict, nextIndex, value
  End Sub

  Public Function Exists(key)
    Exists = ivar_dict.Exists(key)
  End Function

  Public Function Items
    ReDim itemList(ivar_dict.Count - 1)
    Dim i
    For i = 0 To ivar_dict.Count - 1
      BindAt itemList, i, ivar_dict(i)
    Next
    Items = itemList
  End Function

  Public Function Keys
    ReDim keyList(ivar_dict.Count - 1)
    Dim i
    For i = 0 To ivar_dict.Count - 1
      keyList(i) = i
    Next
    Keys = keyList
  End Function

  Public Sub RemoveAll
    ivar_dict.RemoveAll
  End Sub
End Class

Dim ShowValue_Quote
Set ShowValue_Quote = New RegExp
ShowValue_Quote.Pattern = """"
ShowValue_Quote.Global = True

Function ShowValue(value)
  Dim r
  If VarType(value) = vbString Then
    r = """" & ShowValue_Quote.Replace(value, """""") & """"
  ElseIf IsArray(value) Then
    Dim i, sep: sep = ""
    r = "["
    For Each i In value
      r = r & sep & ShowValue(i)
      sep = ","
    Next
    r = r & "]"
  ElseIf IsObject(value) Then
    r = "<" & TypeName(value) & ">"
  ElseIf IsEmpty(value) Then
    r = "<empty>"
  ElseIf IsNull(value) Then
    r = "<null>"
  Else
    On Error Resume Next
    r = CStr(value)
    If Err.Number <> 0 Then
      Err.Clear
      r = "<unknown:" & VarType(value) & ">"
    End If
    On Error GoTo 0
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
  className = "ObjectProperty_Accessor_" & name
  classExpr = "Class " & className & vbNewLine & _
              "  Public Default Property Get Prop(obj)" & vbNewLine & _
              "    Bind Prop, obj." & name & vbNewLine & _
              "  End Property" & vbNewLine & _
              "" & vbNewLine & _
              "  Public Property Let Prop(obj, value)" & vbNewLine & _
              "    obj." & name & " = value" & vbNewLine & _
              "  End Property" & vbNewLine & _
              "" & vbNewLine & _
              "  Public Property Set Prop(obj, value)" & vbNewLine & _
              "    Set obj." & name & " = value" & vbNewLine & _
              "  End Property" & vbNewLine & _
              "End Class" & vbNewLine
  ExecuteGlobal classExpr
  Set ObjectProperty_CreateAccessor = Eval("New " & className)
End Function

Function ObjectProperty_GetAccessor(Name)
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

Sub Sort(list, compare)
  HeapSort list, compare
End Sub

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

Class MessageWriter
  Private ivar_buffer
  
  Private Sub Class_Initialize
    Set ivar_buffer = New ListBuffer
  End Sub

  Public Sub Write(message)
    ivar_buffer.Add message
  End Sub

  Public Default Sub WriteLine(message)
    Write message & vbNewLine
  End Sub

  Private Function FlushBuffer
    Dim s, msg
    s = ""
    For Each msg In ivar_buffer.Items
      s = s & msg
    Next
    ivar_buffer.RemoveAll
    FlushBuffer = s
  End Function

  Private Sub PopupMessage(message)
    MsgBox message, vbOKOnly + vbInformation, WScript.ScriptName
  End Sub

  Public Sub Flush
    Dim s
    s = FlushBuffer

    On Error Resume Next
    WScript.StdOut.Write s

    If Err.Number <> 0 Then
      Err.Clear
      On Error GoTo 0
      PopupMessage s
    End if
  End Sub

  Public Sub FlushAndWriteLine(lastMessage)
    Write(lastMessage)

    Dim s
    s = FlushBuffer

    On Error Resume Next
    WScript.StdOut.Write s

    If Err.Number <> 0 Then
      Err.Clear
      On Error GoTo 0
      PopupMessage s
    Else
      On Error GoTo 0
      WScript.StdOut.WriteLine
    End if
  End Sub
End Class

Dim MsgOut
Set MsgOut = New MessageWriter


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
