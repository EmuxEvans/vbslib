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

Class UnitTest_Assertion
  Public Sub Assert(result)
    AssertWithMessage result, Empty
  End Sub

  Public Sub AssertWithMessage(result, message)
    If Not result Then
      Dim errMsg
      errMsg = "Assert NG."
      If Not IsEmpty(message) Then
        errMsg = errMsg & " [" & message & "]"
      End If
      Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg
    End If
  End Sub

  Public Sub AssertEqual(expected, actual)
    AssertEqualWithMessage expected, actual, Empty
  End Sub

  Public Sub AssertEqualWithMessage(expected, actual, message)
    If expected <> actual Then
      Dim errMsg
      errMsg = "AssertEqual NG: expected <" & ShowValue(expected) & "> but was <" & ShowValue(actual) & ">."
      If Not IsEmpty(message) Then
        errMsg = errMsg & " [" & message & "]"
      End If
      Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg
    End If
  End Sub

  Public Sub AssertSame(expected, actual)
    AssertSameWithMessage expected, actual, Empty
  End Sub

  Public Sub AssertSameWithMessage(expected, actual, message)
    If Not actual Is expected Then
      Dim errMsg
      errMsg = "AssertSame NG: expected <" & TypeName(expected) & "> but was <" & TypeName(actual) & ">."
      If Not IsEmpty(message) Then
        errMsg = errMsg & "[" & message & "]"
      End If
      Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg
    End If
  End Sub

  Public Sub AssertMatch(pattern, text)
    AssertMatchWithMessage pattern, text, Empty
  End Sub

  Public Sub AssertMatchWithMessage(pattern, text, message)
    Dim re
    If IsObject(pattern) Then
      Set re = pattern
    Else
      Set re = New RegExp
      re.Pattern = pattern
    End If

    If Not re.Test(text) Then
      Dim errMsg
      errMsg = "AssertMatch NG: <" & text & "> expected to be match <" & pattern & ">."
      If Not IsEmpty(message) Then
        errMsg = errMsg & "[" & message & "]"
      End If
      Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg
    End If
  End Sub

  Public Sub AssertFail
    AssertFailWithMessage Empty
  End Sub

  Public Sub AssertFailWithMessage(message)
    Dim errMsg
    errMsg = "AssertFail NG."
    If Not IsEmpty(message) Then
      errMsg = errMsg & "[" & message & "]"
    End If
    Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg
  End Sub
End Class

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
    ivar_scriptControl.AddObject "__UnitTest_Assertion__", New UnitTest_Assertion, True
  End Sub

  Public Sub AddObject(name, object)
    ivar_scriptControl.AddObject name, object
  End Sub

  Public Sub ImportTestCase(path)
    ivar_scriptControl.Modules.Add path
    
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
