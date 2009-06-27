' stdlib.vbs: Object Method test.
' @import ../lib/stdlib.vbs

Option Explicit

Class Foo
  Private ivar_method0_countCall
  Private ivar_method1_countCall
  Private ivar_method2_countCall
  Private ivar_func0_countCall
  Private ivar_func1_countCall
  Private ivar_func2_countCall

  Private Sub Class_Initialize
    ivar_method0_countCall = 0
    ivar_method1_countCall = 0
    ivar_method2_countCall = 0
    ivar_func0_countCall = 0
    ivar_func1_countCall = 0
    ivar_func2_countCall = 0
  End Sub

  Public Sub Method0
    ivar_method0_countCall = ivar_method0_countCall + 1
  End Sub

  Public Property Get Method0_CountCall
    Method0_CountCall = ivar_method0_countCall
  End Property

  Public Sub Method1(arg1)
    ivar_method1_countCall = ivar_method1_countCall + 1
  End Sub

  Public Property Get Method1_CountCall
    Method1_CountCall = ivar_method1_countCall
  End Property

  Public Sub Method2(arg1, arg2)
    ivar_method2_countCall = ivar_method2_countCall + 1
  End Sub

  Public Property Get Method2_CountCall
    Method2_CountCall = ivar_method2_countCall
  End Property

  Public Function Func0
    ivar_func0_countCall = ivar_func0_countCall + 1
    Func0 = "Func0"
  End Function

  Public Property Get Func0_CountCall
    Func0_CountCall = ivar_func0_countCall
  End Property

  Public Function Func1(arg1)
    ivar_func1_countCall = ivar_func1_countCall + 1
    Func1 = "Func1:" & arg1
  End Function

  Public Property Get Func1_CountCall
    Func1_CountCall = ivar_func1_countCall
  End Property

  Public Function Func2(arg1, arg2)
    ivar_func2_countCall = ivar_func2_countCall + 1
    Func2 = "Func2:" & Join(Array(arg1, arg2), ",")
  End Function

  Public Property Get Func2_CountCall
    Func2_CountCall = ivar_func2_countCall
  End Property
End Class

Dim obj_foo

Sub SetUp
  Set obj_foo = New Foo
End Sub

Sub TearDown
  Set obj_foo = Nothing
End Sub

Sub TestExecObjectMethodFuncProc0
  ExecObjectMethodFuncProc obj_foo, "Method0", Array()
  AssertEqual 1, obj_foo.Method0_CountCall
End Sub

Sub TestExecObjectMethodFuncProc0_manyCall
  Dim i
  For i = 1 To 100
    ExecObjectMethodFuncProc obj_foo, "Method0", Array()
  Next
  AssertEqual 100, obj_foo.Method0_CountCall
End Sub

Sub TestExecObjectMethodFuncProc1
  ExecObjectMethodFuncProc obj_foo, "Method1", Array("a")
  AssertEqual 1, obj_foo.Method1_CountCall
End Sub

Sub TestExecObjectMethodFuncProc1_manyCall
  Dim i
  For i = 1 To 100
    ExecObjectMethodFuncProc obj_foo, "Method1", Array("a")
  Next
  AssertEqual 100, obj_foo.Method1_CountCall
End Sub

Sub TestExecObjectMethodFuncProc2
  ExecObjectMethodFuncProc obj_foo, "Method2", Array("a", "b")
  AssertEqual 1, obj_foo.Method2_CountCall
End Sub

Sub TestExecObjectMethodFuncProc2_manyCall
  Dim i
  For i = 1 To 100
    ExecObjectMethodFuncProc obj_foo, "Method2", Array("a", "b")
  Next
  AssertEqual 100, obj_foo.Method2_CountCall
End Sub

Sub TestExecObjectMethodFuncProc0
  AssertEqual "Func0", ExecObjectMethodFuncProc(obj_foo, "Func0", Array())
  AssertEqual 1, obj_foo.Func0_CountCall
End Sub

Sub TestExecObjectMethodFuncProc0_manyCall
  Dim i
  For i = 1 To 100
    AssertEqual "Func0", ExecObjectMethodFuncProc(obj_foo, "Func0", Array())
  Next
  AssertEqual 100, obj_foo.Func0_CountCall
End Sub

Sub TestExecObjectMethodFuncProc1
  AssertEqual "Func1:a", ExecObjectMethodFuncProc(obj_foo, "Func1", Array("a"))
  AssertEqual 1, obj_foo.Func1_CountCall
End Sub

Sub TestExecObjectMethodFuncProc1_manyCall
  Dim i
  For i = 1 To 100
    AssertEqual "Func1:a", ExecObjectMethodFuncProc(obj_foo, "Func1", Array("a"))
  Next
  AssertEqual 100, obj_foo.Func1_CountCall
End Sub

Sub TestExecObjectMethodFuncProc2
  AssertEqual "Func2:a,b", ExecObjectMethodFuncProc(obj_foo, "Func2", Array("a", "b"))
  AssertEqual 1, obj_foo.Func2_CountCall
End Sub

Sub TestExecObjectMethodFuncProc2_manyCall
  Dim i
  For i = 1 To 100
    AssertEqual "Func2:a,b", ExecObjectMethodFuncProc(obj_foo, "Func2", Array("a", "b"))
  Next
  AssertEqual 100, obj_foo.Func2_CountCall
End Sub

Sub TestGetObjectMethodSubProc0
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "Method0", 0)
  proc
  AssertEqual 1, obj_foo.Method0_CountCall
End Sub

Sub TestGetObjectMethodSubProc0_manyCall
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "Method0", 0)
  Dim i
  For i = 1 To 100
    proc
  Next
  AssertEqual 100, obj_foo.Method0_CountCall
End Sub

Sub TestGetObjectMethodSubProc0_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodSubProc(obj_foo, "Method0", 0)
    proc
  Next
  AssertEqual 100, obj_foo.Method0_CountCall
End Sub

Sub TestGetObjectMethodSubProc1
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "Method1", 1)
  proc "a"
  AssertEqual 1, obj_foo.Method1_CountCall
End Sub

Sub TestGetObjectMethodSubProc1_manyCall
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "Method1", 1)
  Dim i
  For i = 1 To 100
    proc "a"
  Next
  AssertEqual 100, obj_foo.Method1_CountCall
End Sub

Sub TestGetObjectMethodSubProc1_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodSubProc(obj_foo, "Method1", 1)
    proc "a"
  Next
  AssertEqual 100, obj_foo.Method1_CountCall
End Sub

Sub TestGetObjectMethodSubProc2
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "Method2", 2)
  proc "a", "b"
  AssertEqual 1, obj_foo.Method2_CountCall
End Sub

Sub TestGetObjectMethodSubProc2_manyCall
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "Method2", 2)
  Dim i
  For i = 1 To 100
    proc "a", "b"
  Next
  AssertEqual 100, obj_foo.Method2_CountCall
End Sub

Sub TestGetObjectMethodSubProc2_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodSubProc(obj_foo, "Method2", 2)
    proc "a", "b"
  Next
  AssertEqual 100, obj_foo.Method2_CountCall
End Sub

Sub TestGetObjectMethodFuncProc0
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "Func0", 0)
  AssertEqual "Func0", proc
  AssertEqual 1, obj_foo.Func0_CountCall
End Sub

Sub TestGetObjectMethodFuncProc0_manyCall
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "Func0", 0)
  Dim i
  For i = 1 To 100
    AssertEqual "Func0", proc
  Next
  AssertEqual 100, obj_foo.Func0_CountCall
End Sub

Sub TestGetObjectMethodFuncProc0_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodFuncProc(obj_foo, "Func0", 0)
    AssertEqual "Func0", proc
  Next
  AssertEqual 100, obj_foo.Func0_CountCall
End Sub

Sub TestGetObjectMethodFuncProc1
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "Func1", 1)
  AssertEqual "Func1:a", proc("a")
  AssertEqual 1, obj_foo.Func1_CountCall
End Sub

Sub TestGetObjectMethodFuncProc1_manyCall
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "Func1", 1)
  Dim i
  For i = 1 To 100
    AssertEqual "Func1:a", proc("a")
  Next
  AssertEqual 100, obj_foo.Func1_CountCall
End Sub

Sub TestGetObjectMethodFuncProc1_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodFuncProc(obj_foo, "Func1", 1)
    AssertEqual "Func1:a", proc("a")
  Next
  AssertEqual 100, obj_foo.Func1_CountCall
End Sub

Sub TestGetObjectMethodFuncProc2
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "Func2", 2)
  AssertEqual "Func2:a,b", proc("a", "b")
  AssertEqual 1, obj_foo.Func2_CountCall
End Sub

Sub TestGetObjectMethodFuncProc2_manyCall
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "Func2", 2)
  Dim i
  For i = 1 To 100
    AssertEqual "Func2:a,b", proc("a", "b")
  Next
  AssertEqual 100, obj_foo.Func2_CountCall
End Sub

Sub TestGetObjectMethodFuncProc2_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodFuncProc(obj_foo, "Func2", 2)
    AssertEqual "Func2:a,b", proc("a", "b")
  Next
  AssertEqual 100, obj_foo.Func2_CountCall
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
