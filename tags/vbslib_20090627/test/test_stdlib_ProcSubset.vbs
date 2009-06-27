' stdlib.vbs: Procedure subset test.
' @import ../lib/stdlib.vbs

Option Explicit

Dim Foo_Result

Sub Foo(a, b, c, d, e)
  Foo_Result = Join(Array(a, b, c, d, e), ",")
End Sub

Function Bar(a, b, c, d, e)
  Bar = Join(Array(a, b, c, d, e), ",")
End Function

Sub SetUp
  Foo_Result = Empty
End Sub

Sub TestGetSubProcSubset_ParamsArray
  Dim proc
  Set proc = GetSubProcSubset(GetRef("Foo"), 5, Array("foo", "bar", "baz"))

  proc 0, 1
  AssertEqual "foo,bar,baz,0,1", Foo_Result

  proc 2, 3
  AssertEqual "foo,bar,baz,2,3", Foo_Result
End Sub

Sub TestGetSubProcSubset_ParamsDict
  Dim proc
  Set proc = GetSubProcSubset(GetRef("Foo"), 5, D(Array(0, "foo", 2, "bar", 4, "baz")))

  proc 0, 1
  AssertEqual "foo,0,bar,1,baz", Foo_Result

  proc 2, 3
  AssertEqual "foo,2,bar,3,baz", Foo_Result
End Sub

Sub TestFuncProcSubset_ParamsArray
  Dim proc
  Set proc = GetFuncProcSubset(GetRef("Bar"), 5, Array("foo", "bar", "baz"))
  AssertEqual "foo,bar,baz,0,1", proc(0, 1)
  AssertEqual "foo,bar,baz,2,3", proc(2, 3)
End Sub

Sub TestFuncProcSubset_ParamsDict
  Dim proc
  Set proc = GetFuncProcSubset(GetRef("Bar"), 5, D(Array(0, "foo", 2, "bar", 4, "baz")))
  AssertEqual "foo,0,bar,1,baz", proc(0, 1)
  AssertEqual "foo,2,bar,3,baz", proc(2, 3)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
