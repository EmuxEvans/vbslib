'WMI library

Option Explicit

'=========================================================
'################ common utilities: begin ################
'---------------------------------------------------------

Sub Bind(toStore, value)
  If IsObject(value) Then
    Set toStore = value
  Else
    toStore = value
  End
End Sub

Sub BindAt(toStore, index, value)
  If IsObject(value) Then
    Set toStore(index) = value
  Else
    toStore(index) = value
  End
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
      Err.Raise 9, "wmilib.vbs:ListBuffer.Item(Get)", "out of range."
    End
  End Property

  Public Default Property Let Item(index, value)
    If ivar_dict.Exists(index) Then
      ivar_dict(index) = value
    Else
      Err.Raise 9, "wmilib.vbs:ListBuffer.Item(Let)", "out of range."
    End If
  End Property

  Public Default Property Set Item(index, value)
    If ivar_dict.Exists(index) Then
      Set ivar_dict(index) = value
    Else
      Err.Raise 9, "wmilib.vbs:ListBuffer.Item(Let)", "out of range."
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
      BindAt keyList, i, ivar_dict(i)
    Next
    Keys = keyList
  End Function

  Public Sub RemoveAll
    ivar_dict.RemoveAll
  End Sub
End Class

'---------------------------------------------------------
'################# common utilities: end #################
'=========================================================

'=======================================================
'################ sort utilities: begin ################
'-------------------------------------------------------

Sub SwapArrayElement(list, i, j)
  Dim t
  Bind t, list(i)
  BindAt list, i, list(j)
  BindAt list, j, t
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
      SwapArrayElement list, nextIndex, i
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
    SwapArrayElement list, 0, i
    DownHeap list, 0, i - 1, compare
  Next
End Sub

Sub Sort(list, compare)
  HeapSort list, compare
End Sub

Class NumberCompare
  Public Default Function Compare(a, b)
    Compare = a - b
  End Function
End Class

Class StringTextCompare
  Public Default Function Compare(a, b)
    Compare = StrComp(a, b, vbTextCompare)
  End Function
End Class

Class StringBinaryCompare
  Public Default Function Compare(a, b)
    Compare = StrComp(a, b, vbBinaryCompare)
  End Function
End Class

Class ObjectPropertyCompare
  Private propName
  Private propComp

  Public Property Let PropertyName(value)
    propName = value
  End Property

  Public Property Set PropertyCompare(value)
    Set propComp = value
  End Property

  Public Default Function Compare(a, b)
    If IsEmpty(propName) Then
      Err.Raise 51, "wmilib.vbs:ObjectPropertyCompare", "Not defined `PropertyName'."
    End If
    If IsEmpty(propComp) Then
      Err.Raise 51, "wmilib.vbs:ObjectPropertyCompare", "Not defined `PropertyCompare'."
    End If
    Compare = propComp(GetObjectProperty(a, propName), _
                       GetObjectProperty(b, propName))
  End Function
End Class

Function New_ObjectPropertyCompare(propertyName, propertyCompare)
  Dim compare
  Set compare = New ObjectPropertyCompare
  compare.PropertyName = propertyName
  Set compare.PropertyCompare = propertyCompare
  Set New_ObjectPropertyCompare = compare
End Function

'-------------------------------------------------------
'################# sort utilities: end #################
'=======================================================

'======================================================
'################ I/O utilities: begin ################
'------------------------------------------------------

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
      On Error ToTo 0
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

'------------------------------------------------------
'################# I/O utilities: end #################
'======================================================

'=======================================================
'################# MOP utilities: Begin ################
'-------------------------------------------------------

Function ExistsObjectProperty(Object, name)
  Dim value
  On Error Resume Next
  value = Eval("VarType(object." & name & ")")
  Select Case Err.Number
    Case 0:
      ExistsObjectProperty = True
    Case 438:
      ExistsObjectProperty = False
    Case Else:
      Err.Raise 51, "wmilib.vbs:ExistsObjectProperty", "unknown error."
  End Select
end Function

Function GetObjectProperty(object, name)
  Bind GetObjectProperty, Eval("object." & Name)
End Function

Function ShowObjectProperty(object, name)
  Dim value
  value = GetObjectProperty(object, propertyName)
  ShowObjectProperty = propertyName & ": " & value
End Function

'-------------------------------------------------------
'################## MOP utilities: end #################
'=======================================================

Function WMIServiceInstancesOf(computerName, serviceName)
  Dim wbemServices
  Set wbemServices = GetObject("winmgmts:\\" & computerName)
  Set WMIServiceInstancesOf = wbemServices.InstancesOf(serviceName)
End Function

Function DefinedObjectProperty(object, propertyName)
  Dim value
  On Error Resume Next
  value = Eval("VarType(object." & propertyName & ")")
  Select Case Err.Number
    Case 0:
      DefinedObjectProperty = True
    Case 438:
      DefinedObjectProperty = False
    Case Else:
      Err.Raise Err.Number
  End Select
End Function



' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
