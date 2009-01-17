'Read Eval Print Loop

Option Explicit

Const MAX_HISTORY = 30

Class History
  Private dict
  Private firstIndex
  Private lastIndex
  Private maxHistory

  Public Sub Class_Initialize
    Set dict = CreateObject("Scripting.Dictionary")
    firstIndex = 0
    lastIndex = 0
    maxHistory = MAX_HISTORY
  End Sub

  Public Property Get NextIndex
    NextIndex = lastIndex
  End Property

  Public Sub Add(expr)
    dict(lastIndex) = expr
    lastIndex = lastIndex + 1

    Do While dict.Count > maxHistory
      dict.Remove firstIndex
      firstIndex = firstIndex + 1
    Loop
  End Sub

  Public Default Property Get Item(index)
    If dict.Exists(index) Then
      Item = dict(index)
    End If
  End Property

  Public Function Exists(index)
    Exists = dict.Exists(index)
  End Function

  Public Function Keys
    ReDim KeyList(dict.Count - 1)

    Dim i
    For i = 0 To dict.Count - 1
      KeyList(i) = firstIndex + i
    Next

    Keys = KeyList
  End Function
End Class

Const POPUP_TITLE = "Read Eval Print Loop"

Sub PopupError
  MsgBox Err.Number & ": " & Err.Description & " (" & Err.Source & ")", _
         vbOKOnly + vbCritical, POPUP_TITLE + ": Error"
End Sub

Sub PopupResult(expr, result)
  MsgBox "`" & expr & "'" & vbNewLine & "=> " & result, _
         vbOKOnly, POPUP_TITLE & ": Result"
End Sub

Sub PopupHistory(hist)
  Dim i, text, sep
  For Each i In hist.Keys
    text = text & sep & i & ": " & hist(i)
    sep = vbNewLine
  Next
  MsgBox text, vbOKOnly + vbInformation, POPUP_TITLE & ": History"
End Sub

Sub REPL_Execute(expr)
  On Error Resume Next
  ExecuteGlobal expr

  If Err.Number <> 0 Then
    PopupError
  End If

  Err.Clear
End Sub

Sub REPL_Evaluate(expr)
  Dim result
  On Error Resume Next
  result = Eval(expr)

  If Err.Number = 0 Then
    PopupResult expr, result
  Else
    PopupError
  End If

  Err.Clear
End Sub

Dim execCommand
Set execCommand = New RegExp
execCommand.Pattern = "^e *"
execCommand.IgnoreCase = True

Dim evalCommand
Set evalCommand = New RegExp
evalCommand.Pattern = "^p *"
evalCommand.IgnoreCase = True

Dim histCommand
Set histCommand = New RegExp
histCommand.Pattern = "^h$|^h *"
histCommand.IgnoreCase = True

Dim hist
Set hist = New History

Dim expr
Dim defaultExpr
defaultExpr = Empty

Do
  expr = InputBox("input `statement' or `e statement' or `p expression'. `h' for history.", _
                  POPUP_TITLE & " [" & hist.NextIndex & "]", _
                  defaultExpr)

  If IsEmpty(expr) Then
    Exit Do
  End If

  hist.Add expr
  defaultExpr = Empty

  If execCommand.Test(expr) Then
    expr = execCommand.Replace(expr, "")
    REPL_Execute expr
  ElseIf evalCommand.Test(expr) Then
    expr = evalCommand.Replace(expr, "")
    REPL_Evaluate expr
  ElseIf histCommand.Test(expr) Then
    expr = histCommand.Replace(expr, "")
    If expr = "" Then
      PopupHistory(hist)
    Else
      Dim index
      index = CInt(expr)
      If hist.Exists(index) Then
        defaultExpr = hist(index)
      End If
    End If
  Else
    REPL_Execute expr
  End If
Loop

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
