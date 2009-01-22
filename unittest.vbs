' VBScript Unit Test
' need for `stdlib.vbs'

Option Explicit

Const UNITTEST_FAIL_TYPE_E = "Error"
Const UNITTEST_FAIL_TYPE_A = "Assert"
Const UNITTEST_ASSERT_SOURCE_KEYWORD = "UnitTest Assertion"

Dim UnitTest_Desc
UnitTest_Desc = Empty

Dim UnitTest_SetUpSubs
Set UnitTest_SetUpSubs = New ListBuffer

Dim UnitTest_TearDownSubs
Set UnitTest_TearDownSubs = New ListBuffer

Dim UnitTest_TestCaseSubs
Set UnitTest_TestCaseSubs = New ListBuffer

Sub UnitTest_Description(message)
  UnitTest_Desc = message
End Sub

Sub UnitTest_SetUp(setupSub)
  UnitTest_SetUpSubs.Add setupSub
End Sub

Sub UnitTest_TearDown(tearDownSub)
  UnitTest_TearDownSubs.Add tearDownSub
End Sub

Sub UnitTest_TestCase(testCaseSub)
  UnitTest_TestCaseSubs.Add testCaseSub
End Sub

Function UnitTest_MakeErrorEntry(testCaseSub, message)
  ReDim entry(2)
  entry(0) = UNITTEST_FAIL_TYPE_E
  entry(1) = testCaseSub
  entry(2) = "UnitTestError: " & message & ": " & _
             "(" & Err.Number & ") " & "[" & Err.Source & "] " & Err.Description
  UnitTest_MakeErrorEntry = entry
End Function

Function UnitTest_MakeAssertFailEntry(testCaseSub)
  ReDim entry(2)
  entry(0) = UNITTEST_FAIL_TYPE_A
  entry(1) = testCaseSub
  entry(2) = "Assertion Failed: " & Err.Description
  UnitTest_MakeAssertFailEntry = entry
End Function

Sub UnitTest_RunSubs(subs)
  Dim s
  For Each s In subs.Items
    ExecuteGlobal "Call " & s
  Next
End Sub

Sub UnitTest_RunTestCase(testCaseSub, failList)
  On Error Resume Next

  UnitTest_RunSubs UnitTest_SetUpSubs
  If Err.Number = 0 Then
    ExecuteGlobal "Call " & testCaseSub
    If Err.Number <> 0 Then
      If Err.Source = UNITTEST_ASSERT_SOURCE_KEYWORD Then
        failList.Add UnitTest_MakeAssertFailEntry(testCaseSub)
      Else
        failList.Add UnitTest_MakeErrorEntry(testCaseSub, "error test case.")
      End If
      Err.Clear
    End If
  Else
    failList.Add UnitTest_MakeErrorEntry(testCaseSub, "failed to setup.")
    Err.Clear
  End If

  UnitTest_RunSubs UnitTest_TearDownSubs
  If Err.Number <> 0 Then
    failList.Add UnitTest_MakeErrorEntry(testCaseSub, "failed to teardown.")
    Err.Clear
  End If
End Sub

Sub UnitTest_ConsoleRun
  Dim testCaseSub, failEntry, count
  Dim failList: Set failList = New ListBuffer
  Dim allFailList: Set allFailList = New ListBuffer

  If Not IsEmpty(UnitTest_Desc) Then
    WScript.StdOut.WriteLine UnitTest_Desc
  End If

  For Each testCaseSub In UnitTest_TestCaseSubs.Items
    UnitTest_RunTestCase testCaseSub, failList
    If failList.Count = 0 Then
      WScript.StdOut.Write "."
    Else
      WScript.StdOut.Write "E"
    End If
    For Each failEntry In failList.Items
      allFailList.Add failEntry
    Next
    failList.RemoveAll
  Next
  WScript.StdOut.WriteBlankLines 1

  count = 0
  For Each failEntry In allFailList.Items
    count = count + 1
    WScript.StdOut.WriteLine "(" & count & ") [" & failEntry(1) & "] " & failEntry(2)
  Next

  If allFailList.Count = 0 Then
    WScript.Quit 0
  Else
    WScript.StdOut.WriteLine "*"
    WScript.Quit 1
  End If
End Sub

Sub Assert(result)
  AssertWithComment result, Empty
End Sub

Sub AssertWithComment(result, comment)
  If Not result Then
    Dim errMsg
    errMsg = "Assert NG."
    If Not IsEmpty(comment) Then
      errMsg = errMsg & " [" & comment & "]"
    End If
    Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg
  End If
End Sub

Sub AssertEqual(expected, actual)
  AssertEqualWithComment expected, actual, Empty
End Sub

Sub AssertEqualWithComment(expected, actual, comment)
  If expected <> actual Then
    Dim errMsg
    errMsg = "AssertEqual NG: expected <" & expected & "> but was <" & actual & ">."
    If Not IsEmpty(comment) Then
      errMsg = errMsg & " [" & comment & "]"
    End If
    Err.Raise RuntimeError, UNITTEST_ASSERT_SOURCE_KEYWORD, errMsg
  End If
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
