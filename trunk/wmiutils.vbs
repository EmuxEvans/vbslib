'WMI Utilities

Option Explicit

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

Function GetObjectProperty(object, propertyName)
  GetObjectProperty = Eval("object." & propertyName)
End Function

Function ShowObjectProperty(object, propertyName)
  Dim value
  value = GetObjectProperty(object, propertyName)
  ShowObjectProperty = propertyName & ": " & value
End Function

Const MessageWriter_INIT_BUF_SIZE = 15

Class MessageWriter
  Private buffer
  Private lastIndex

  Private Sub Class_Initialize
    ReDim buffer(MessageWriter_INIT_BUF_SIZE)
    lastIndex = 0
  End Sub

  Private Sub ExtendBuffer
    Dim maxIndex
    maxIndex = UBound(buffer)
    If maxIndex < MessageWriter_INIT_BUF_SIZE Then
      maxIndex = MessageWriter_INIT_BUF_SIZE
    Else
      maxIndex = maxIndex * 2
    End If
    ReDim Preserve buffer(maxIndex)
  End Sub

  Public Sub Write(message)
    If lastIndex > UBound(buffer) Then
      ExtendBuffer
    End If
    buffer(lastIndex) = message
    lastIndex = lastIndex + 1
  End Sub

  Public Default Sub WriteLine(message)
    Write message & vbNewLine
  End Sub

  Private Function FlushBuffer
    Dim s
    s = ""

    Dim i
    For i = 0 To lastIndex - 1
      s = s & buffer(i)
      buffer(i) = ""
    Next
    lastIndex = 0

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
      PopupMessage s
    Else
      WScript.StdOut.WriteLine
    End if
  End Sub
End Class

Dim MsgOut
Set MsgOut = New MessageWriter

Sub DownHeap(list, startIndex, maxIndex, compare, swap)
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
      swap list, nextIndex, i
    Else
      Exit Do
    End If

    i = nextIndex
  Loop
End Sub

Sub HeapSort(list, compare, swap)
  Dim i

  For i = Int((UBound(list) - 1) / 2) To 0 Step -1
    DownHeap list, i, UBound(list), compare, swap
  Next

  For i = UBound(list) To 1 Step -1
    swap list, 0, i
    DownHeap list, 0, i - 1, compare, swap
  Next
End Sub

Sub Sort(list, compare, swap)
  HeapSort list, compare, swap
End Sub

Class SwapValue
  Public Default Sub Swap(list, i, j)
    Dim t
    t = list(i)
    list(i) = list(j)
    list(j) = t
  End Sub
End Class

Sub SortValue(list, compare)
  Sort list, compare, New SwapValue
End Sub

Class SwapObject
  Public Default Sub Swap(list, i, j)
    Dim t
    Set t = list(i)
    Set list(i) = list(j)
    Set list(j) = t
  End Sub
End Class

Sub SortObject(list, compare)
  Sort list, compare, New SwapObject
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
      Err.Raise 51, "ObjectPropertyCompare", "Not defined `PropertyName'."
    End If
    If IsEmpty(propComp) Then
      Err.Raise 51, "ObjectPropertyCompare", "Not defined `PropertyCompare'."
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

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End: