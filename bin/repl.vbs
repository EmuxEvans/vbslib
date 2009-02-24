'Read Eval Print Loop

Option Explicit

Const MAX_HISTORY = 30
Const DEFAULT_TIMEOUT_MILLISEC = 4000

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

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Function FileReadAll(path)
  Dim stream
  Set stream = fso.OpenTextFile(path)
  FileReadAll = stream.ReadAll
  stream.Close
End Function

Dim logFilename, logStream
logFilename = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), _
                            fso.GetBaseName(WScript.ScriptFullName) & ".log")
Set logStream = fso.OpenTextFile(logFilename, ForAppending, True)

Const POPUP_TITLE = "Read Eval Print Loop"

Sub PopupMessage(prompt, buttons, title)
  logStream.WriteLine Now
  logStream.WriteLine "[ " & title & " ]"
  logStream.WriteLine prompt
  logStream.WriteBlankLines 1

  MsgBox prompt, buttons, title
End Sub

Function PopupInputBox(prompt, title, default)
  Dim s
  s = InputBox(prompt, title, default)

  logStream.WriteLine Now
  logStream.WriteLine "[ " & title & " ]"
  logStream.WriteLine prompt
  logStream.WriteLine "input: " & ShowValue(s)
  logStream.WriteBlankLines 1

  PopupInputBox = s
End Function

Function PopupFileOpenDialog
  Dim path
  path = InputFileOpenDialog

  logStream.WriteLine Now
  logStream.WriteLine "[ FileOpenDialog ]"
  logStream.WriteLine "GetFilePath: " & ShowValue(path)
  logStream.WriteBlankLines 1

  PopupFileOpenDialog = path
End Function

Sub PopupError(title)
  PopupMessage Err.Number & ": " & Err.Description & " (" & Err.Source & ")", _
               vbOKOnly + vbCritical, POPUP_TITLE + ": " & title
End Sub

Sub PopupResult(expr, result)
  PopupMessage expr & vbNewLine & "=> " & result, _
               vbOKOnly, POPUP_TITLE & ": Result"
End Sub

Sub PopupHistory(hist)
  Dim keys
  keys = hist.Keys

  ReDim histItemList(UBound(keys))
  Dim i
  For i = 0 To UBound(keys)
    histItemList(i) = keys(i) & ": " & hist(keys(i))
  Next

  PopupMessage Join(histItemList, vbNewLine), _
               vbOKOnly + vbInformation, POPUP_TITLE & ": History"
End Sub

Function GetHistory(hist, indexExpr)
  Dim index
  Err.Clear
  On Error Resume Next
  index = CInt(indexExpr)
  If Err.Number = 0 Then
    If hist.Exists(index) Then
      GetHistory = hist(index)
      Exit Function
    End If
  End If
  GetHistory = Empty
End Function

Dim REPL_ScriptControl
Set REPL_ScriptControl = CreateObject("ScriptControl")
REPL_ScriptControl.Language = "VBScript"
REPL_ScriptControl.AddObject "WScript", WScript
REPL_ScriptControl.Timeout = DEFAULT_TIMEOUT_MILLISEC

Sub REPL_Execute(expr)
  Err.Clear
  On Error Resume Next
  REPL_ScriptControl.ExecuteStatement expr
  If Err.Number <> 0 Then
    PopupError("Statement Error")
    Err.Clear
  End If
End Sub

Sub REPL_Evaluate(expr)
  Dim result
  Err.Clear
  On Error Resume Next
  result = ShowValue(REPL_ScriptControl.Eval(expr))
  If Err.Number = 0 Then
    PopupResult expr, result
  Else
    PopupError("Expression Error")
    Err.Clear
  End If
End Sub

Sub PopupCurrentTimeout
  PopupMessage REPL_ScriptControl.Timeout & " milliseconds", _
               vbOKOnly + vbInformation, POPUP_TITLE & ": Current Timeout"
End Sub

Sub SetTimeout(millisec)
  Dim ms
  Err.Clear
  On Error Resume Next
  REPL_ScriptControl.Timeout = CLng(millisec)
  If Err.Number <> 0 Then
    PopupError("Timeout Error")
    Err.Clear
  End If
End Sub

Sub ImportFile(path)
  If IsEmpty(path) Then
    path = PopupFileOpenDialog
  End If
  If Not IsEmpty(path) Then
    Err.Clear
    On Error Resume Next
    REPL_ScriptControl.AddCode FileReadAll(path)
    If Err.Number <> 0 Then
      PopupError("Import Error")
      Err.Clear
    End If
  End If
End Sub

Class PseudoRegexpFilter
  Public Function Test(text)
    Test = True
  End Function
End Class

Sub PopupProcedureList(regexpFilter)
  Dim filter
  If IsEmpty(regexpFilter) Then
    Set filter = New PseudoRegexpFilter
  Else
    Set filter = New RegExp
    filter.Pattern = regexpFilter
    filter.IgnoreCase = True
  End If

  REPL_ScriptControl.AddCode "Now"      ' dummy AddCode to update ScriptControl.Procedures
  Dim procSet
  Set procSet = REPL_ScriptControl.Procedures

  ReDim procItemList(procSet.Count - 1)
  Dim count, proc, i, s, sep
  count = 0
  For Each proc In procSet
    If filter.Test(proc.Name) Then
      s = (count + 1) & ". "
      If proc.HasReturnValue Then
        s = s & "Function "
      Else
        s = s & "Sub "
      End If
      s = s & proc.Name & "("
      sep = ""
      For i = 1 To proc.NumArgs
        s = s & sep & "a" & i
        sep = ","
      Next
      s = s & ")"
      procItemList(count) = s
      count = count + 1
    End If
  Next
  ReDim Preserve procItemList(count - 1)

  PopupMessage Join(procItemList, vbNewLine), _
               vbOKOnly + vbInformation, POPUP_TITLE & ": Defined Procedures"
End Sub

Sub ScriptEngineReset
  REPL_ScriptControl.Reset
End Sub

Sub PopupHelp
  PopupMessage Join(Array("Statement", _
                          "e Statement", _
                          "p Expression", _
                          "h", _
                          "h Index", _
                          "hh", _
                          "@timeout", _
                          "@timeout MilliSeconds", _
                          "@import", _
                          "@import VBScriptFile", _
                          "@proc", _
                          "@proc NamePattern", _
                          "@Reset", _
                          "?"), _
                    vbNewLine), _
               vbOKOnly + vbInformation, POPUP_TITLE & ": Help"
End Sub

Dim execCommand
Set execCommand = New RegExp
execCommand.Pattern = "^e\s+"
execCommand.IgnoreCase = True

Dim evalCommand
Set evalCommand = New RegExp
evalCommand.Pattern = "^p\s+"
evalCommand.IgnoreCase = True

Dim histCommand
Set histCommand = New RegExp
histCommand.Pattern = "^h$|^hh$|^h\s+"
histCommand.IgnoreCase = True

Dim timeoutCommand
Set timeoutCommand = New RegExp
timeoutCommand.Pattern = "^@timeout$|^@timeout\s+"
timeoutCommand.IgnoreCase = True

Dim importCommand
Set importCommand = New RegExp
importCommand.Pattern = "^@import$|^@import\s+"
importCommand.IgnoreCase = True

Dim procCommand
Set procCommand = New RegExp
procCommand.Pattern = "^@proc$|^@proc\s+"
procCommand.IgnoreCase = True

Dim resetCommand
Set resetCommand = New RegExp
resetCommand.Pattern = "^@reset$"
resetCommand.IgnoreCase = True

Dim helpCommand
Set helpCommand = New RegExp
helpCommand.Pattern = "^\?$"
helpCommand.IgnoreCase = True

Dim hist
Set hist = New History

Dim expr
Dim defaultExpr
defaultExpr = Empty

Do
  expr = PopupInputBox("Input `statement' or `e statement' or `p expression'. `h' for history. `?' for help.", _
                       POPUP_TITLE & " [" & hist.NextIndex & "]", _
                       defaultExpr)

  If IsEmpty(expr) Then
    Exit Do
  End If

  hist.Add expr
  defaultExpr = Empty

  If execCommand.Test(expr) Then
    REPL_Execute execCommand.Replace(expr, "")
  ElseIf evalCommand.Test(expr) Then
    REPL_Evaluate evalCommand.Replace(expr, "")
  ElseIf histCommand.Test(expr) Then
    Select Case LCase(expr)
      Case "h":
        PopupHistory(hist)
      Case "hh":
        defaultExpr = hist(hist.NextIndex - 2)
      Case Else:
        defaultExpr = GetHistory(hist, histCommand.Replace(expr, ""))
    End Select
  ElseIf timeoutCommand.Test(expr) Then
    Select Case LCase(expr)
      Case "@timeout":
        PopupCurrentTimeout
      Case Else:
        SetTimeout timeoutCommand.Replace(expr, "")
    End Select
  ElseIf importCommand.Test(expr) Then
    Select Case LCase(expr)
      Case "@import":
        ImportFile Empty
      Case Else:
        expr = importCommand.Replace(expr, "")
        ImportFile expr
    End Select
  ElseIf procCommand.Test(expr) Then
    Select Case LCase(expr)
      Case "@proc":
        PopupProcedureList Empty
      Case Else:
        PopupProcedureList procCommand.Replace(expr, "")
    End Select
  ElseIf resetCommand.Test(expr) Then
    ScriptEngineReset
  ElseIf helpCommand.Test(expr) Then
    PopupHelp
  Else
    REPL_Execute expr
  End If
Loop

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
