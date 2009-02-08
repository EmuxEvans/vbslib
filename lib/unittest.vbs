' VBScript Unit Test
' need for `stdlib.vbs'

Option Explicit

Const UNITTEST_ASSERT_SOURCE_KEYWORD = "UnitTest Assertion"

Function UnitTest_IsAssertFail(error)
  If error.Source = UNITTEST_ASSERT_SOURCE_KEYWORD Then
    UnitTest_IsAssertFail = True
  Else
    UnitTest_IsAssertFail = False
  End If
End Function

Dim assertExpr
Set assertExpr = New ListBuffer

assertExpr.Add "Sub Assert(result)"
assertExpr.Add "  AssertWithMessage result, Empty"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertWithMessage(result, message)"
assertExpr.Add "  If Not result Then"
assertExpr.Add "    Dim errMsg"
assertExpr.Add "    errMsg = ""Assert NG."""
assertExpr.Add "    If Not IsEmpty(message) Then"
assertExpr.Add "      errMsg = errMsg & "" ["" & message & ""]"""
assertExpr.Add "    End If"
assertExpr.Add "    Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg"
assertExpr.Add "  End If"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertEqual(expected, actual)"
assertExpr.Add "  AssertEqualWithMessage expected, actual, Empty"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertEqualWithMessage(expected, actual, message)"
assertExpr.Add "  If expected <> actual Then"
assertExpr.Add "    Dim errMsg"
assertExpr.Add "    errMsg = ""AssertEqual NG: expected <"" & ShowValue(expected) & ""> but was <"" & ShowValue(actual) & "">."""
assertExpr.Add "    If Not IsEmpty(message) Then"
assertExpr.Add "      errMsg = errMsg & "" ["" & message & ""]"""
assertExpr.Add "    End If"
assertExpr.Add "    Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg"
assertExpr.Add "  End If"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertSame(expected, actual)"
assertExpr.Add "  AssertSameWithMessage expected, actual, Empty"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertSameWithMessage(expected, actual, message)"
assertExpr.Add "  If Not actual Is expected Then"
assertExpr.Add "    Dim errMsg"
assertExpr.Add "    errMsg = ""AssertSame NG: expected <"" & TypeName(expected) & ""> but was <"" & TypeName(actual) & "">."""
assertExpr.Add "    If Not IsEmpty(message) Then"
assertExpr.Add "      errMsg = errMsg & ""["" & message & ""]"""
assertExpr.Add "    End If"
assertExpr.Add "    Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg"
assertExpr.Add "  End If"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertMatch(pattern, text)"
assertExpr.Add "  AssertMatchWithMessage pattern, text, Empty"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertMatchWithMessage(pattern, text, message)"
assertExpr.Add "  Dim re"
assertExpr.Add "  Set re = New RegExp"
assertExpr.Add "  re.Pattern = pattern"
assertExpr.Add "  If Not re.Test(text) Then"
assertExpr.Add "    Dim errMsg"
assertExpr.Add "    errMsg = ""AssertMatch NG: <"" & text & ""> expected to be match <"" & pattern & "">."""
assertExpr.Add "    If Not IsEmpty(message) Then"
assertExpr.Add "      errMsg = errMsg & ""["" & message & ""]"""
assertExpr.Add "    End If"
assertExpr.Add "    Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg"
assertExpr.Add "  End If"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertFail"
assertExpr.Add "  AssertFailWithMessage Empty"
assertExpr.Add "End Sub"
assertExpr.Add ""
assertExpr.Add "Sub AssertFailWithMessage(message)"
assertExpr.Add "  Dim errMsg"
assertExpr.Add "  errMsg = ""AssertFail NG."""
assertExpr.Add "  If Not IsEmpty(message) Then"
assertExpr.Add "    errMsg = errMsg & ""["" & message & ""]"""
assertExpr.Add "  End If"
assertExpr.Add "  Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg"
assertExpr.Add "End Sub"

Dim UnitTest_AssertProcCode
UnitTest_AssertProcCode = Replace(Join(assertExpr.Items, vbNewLine), _
                                  "UNITTEST_ASSERT_SOURCE_KEYWORD", _
                                  """" & UNITTEST_ASSERT_SOURCE_KEYWORD & """")
Set assertExpr = Nothing
ExecuteGlobal UnitTest_AssertProcCode

Dim UnitTest_TestProcConvention
Set UnitTest_TestProcConvention = New RegExp
UnitTest_TestProcConvention.Pattern = "^Test"
UnitTest_TestProcConvention.IgnoreCase = True

Dim UnitTest_ImportAnnotation
Set UnitTest_ImportAnnotation = New RegExp
UnitTest_ImportAnnotation.Pattern = "^'\s*@import\s+(\S|\S.*\S)\s*$"
UnitTest_ImportAnnotation.IgnoreCase = True
UnitTest_ImportAnnotation.Global = True
UnitTest_ImportAnnotation.Multiline = True

Class UnitTest_TestProc
  Private ivar_testModule
  Private ivar_procName
  Private ivar_hasSetUp
  Private ivar_hasTearDown

  Public Sub Build(testModule, procName, hasSetUp, hasTearDown)
    Set ivar_testModule = testModule
    ivar_procName = procName
    ivar_hasSetUp = hasSetUp
    ivar_hasTearDown = hasTearDown
  End Sub

  Public Property Get ModuleName
    ModuleName = ivar_testModule.Name
  End Property

  Public Property Get Name
    Name = ivar_procName
  End Property

  Public Sub SetUp
    If ivar_hasSetUp Then
      ivar_testModule.Run "SetUp"
    End If
  End Sub

  Public Sub TearDown
    If ivar_hasTearDown Then
      ivar_testModule.Run "TearDown"
    End If
  End Sub

  Public Sub Execute
    ivar_testModule.Run ivar_procName
  End Sub
End Class

Class UnitTest_TestCase
  Private ivar_testModule

  Public Sub Build(testModule)
    Set ivar_testModule = testModule
  End Sub

  Public Function HasSetUp
    Dim proc
    For Each proc In ivar_testModule.Procedures
      If UCase(proc.Name) = "SETUP" Then
        HasSetUp = True
        Exit Function
      End If
    Next
    HasSetUp = False
  End Function

  Public Function HasTearDown
    Dim proc
    For Each proc In ivar_testModule.Procedures
      If UCase(proc.Name) = "TEARDOWN" Then
        HasTearDown = True
        Exit Function
      End If
    Next
    HasTearDown = False
  End Function

  Public Function Items
    Dim hasSetUp_, hasTearDown_
    hasSetUp_ = HasSetUp
    hasTearDown_ = HasTearDown

    Dim procList, proc
    Set procList = New ListBuffer
    For Each proc In ivar_testModule.Procedures
      If UnitTest_TestProcConvention.Test(proc.Name) Then
        procList.Add New UnitTest_TestProc
        procList.LastItem.Build ivar_testModule, proc.Name, hasSetUp_, hasTearDown_
      End If
    Next

    Items = procList.Items
  End Function
End Class

Class UnitTest_TestCaseLoader
  Private ivar_fso
  Private ivar_scriptControl

  Private Sub Class_Initialize
    Set ivar_fso = CreateObject("Scripting.FileSystemObject")
    Set ivar_scriptControl = CreateObject("ScriptControl")
    ivar_scriptControl.Language = "VBScript"
  End Sub

  Public Sub AddObject(name, object)
    ivar_scriptControl.AddObject name, object
  End Sub

  Public Sub ImportTestCase(path)
    ivar_scriptControl.Modules.Add path
    ivar_scriptControl.Modules(path).AddCode UnitTest_AssertProcCode
    
    Dim stream, code
    Set stream = ivar_fso.OpenTextFile(path)
    code = stream.ReadAll
    stream.Close

    Dim match, libPath
    For Each match In UnitTest_ImportAnnotation.Execute(code)
      libPath = match.SubMatches(0)
      Set stream = ivar_fso.OpenTextFile(libPath)
      ivar_scriptControl.Modules(path).AddCode stream.ReadAll
      stream.Close
    Next

    ivar_scriptControl.Modules(path).AddCode code
  End Sub

  Public Function Items
    Dim modList, mo
    Set modList = New ListBuffer
    For Each mo In ivar_scriptControl.Modules
      If mo.Name <> "Global" Then
        modList.Add New UnitTest_TestCase
        modList.LastItem.Build mo
      End If
    Next

    Items = modList.Items
  End Function
End Class

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
