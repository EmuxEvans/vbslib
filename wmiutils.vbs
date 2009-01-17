'WMI Utilities

Option Explicit

Function WMIServiceInstancesOf(computerName, serviceName)
  Dim wbemServices
  Set wbemServices = GetObject("winmgmts:\\" & computerName)
  Set WMIServiceInstancesOf = wbemServices.InstancesOf(serviceName)
End Function

Const MessageWriter_INIT_BUF_SIZE = 15

Class MessageWriter
  Private buffer
  Private lastIndex

  Private Sub Class_Initialize
    ReDim buffer(MessageWriter_INIT_BUF_SIZE)
    lastIndex = 0
  End Sub

  Private Sub ExpandBuffer
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
      ExpandBuffer
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

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
