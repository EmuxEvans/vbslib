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

  Public Sub method0
    ivar_method0_countCall = ivar_method0_countCall + 1
  End Sub

  Public Property Get method0_countCall
    method0_countCall = ivar_method0_countCall
  End Property

  Public Sub method1(arg1)
    ivar_method1_countCall = ivar_method1_countCall + 1
  End Sub

  Public Property Get method1_countCall
    method1_countCall = ivar_method1_countCall
  End Property

  Public Sub method2(arg1, arg2)
    ivar_method2_countCall = ivar_method2_countCall + 1
  End Sub

  Public Property Get method2_countCall
    method2_countCall = ivar_method2_countCall
  End Property

  Public Function func0
    ivar_func0_countCall = ivar_func0_countCall + 1
    func0 = "func0"
  End Function

  Public Property Get func0_countCall
    func0_countCall = ivar_func0_countCall
  End Property

  Public Function func1(arg1)
    ivar_func1_countCall = ivar_func1_countCall + 1
    func1 = "func1:" & arg1
  End Function

  Public Property Get func1_countCall
    func1_countCall = ivar_func1_countCall
  End Property

  Public Function func2(arg1, arg2)
    ivar_func2_countCall = ivar_func2_countCall + 1
    func2 = "func2:" & Join(Array(arg1, arg2), ",")
  End Function

  Public Property Get func2_countCall
    func2_countCall = ivar_func2_countCall
  End Property
End Class

Dim obj_foo

Sub SetUp
  Set obj_foo = New Foo
End Sub

Sub TearDown
  Set obj_foo = Nothing
End Sub

Sub TestInvokeObjectMethod0
  InvokeObjectMethod obj_foo, "method0", Array()
  AssertEqual 1, obj_foo.method0_countCall
End Sub

Sub TestInvokeObjectMethod0_manyCall
  Dim i
  For i = 1 To 100
    InvokeObjectMethod obj_foo, "method0", Array()
  Next
  AssertEqual 100, obj_foo.method0_countCall
End Sub

Sub TestInvokeObjectMethod1
  InvokeObjectMethod obj_foo, "method1", Array("a")
  AssertEqual 1, obj_foo.method1_countCall
End Sub

Sub TestInvokeObjectMethod1_manyCall
  Dim i
  For i = 1 To 100
    InvokeObjectMethod obj_foo, "method1", Array("a")
  Next
  AssertEqual 100, obj_foo.method1_countCall
End Sub

Sub TestInvokeObjectMethod2
  InvokeObjectMethod obj_foo, "method2", Array("a", "b")
  AssertEqual 1, obj_foo.method2_countCall
End Sub

Sub TestInvokeObjectMethod2_manyCall
  Dim i
  For i = 1 To 100
    InvokeObjectMethod obj_foo, "method2", Array("a", "b")
  Next
  AssertEqual 100, obj_foo.method2_countCall
End Sub

Sub TestFuncallObjectMethod0
  AssertEqual "func0", FuncallObjectMethod(obj_foo, "func0", Array())
  AssertEqual 1, obj_foo.func0_countCall
End Sub

Sub TestFuncallObjectMethod0_manyCall
  Dim i
  For i = 1 To 100
    AssertEqual "func0", FuncallObjectMethod(obj_foo, "func0", Array())
  Next
  AssertEqual 100, obj_foo.func0_countCall
End Sub

Sub TestFuncallObjectMethod1
  AssertEqual "func1:a", FuncallObjectMethod(obj_foo, "func1", Array("a"))
  AssertEqual 1, obj_foo.func1_countCall
End Sub

Sub TestFuncallObjectMethod1_manyCall
  Dim i
  For i = 1 To 100
    AssertEqual "func1:a", FuncallObjectMethod(obj_foo, "func1", Array("a"))
  Next
  AssertEqual 100, obj_foo.func1_countCall
End Sub

Sub TestFuncallObjectMethod2
  AssertEqual "func2:a,b", FuncallObjectMethod(obj_foo, "func2", Array("a", "b"))
  AssertEqual 1, obj_foo.func2_countCall
End Sub

Sub TestFuncallObjectMethod2_manyCall
  Dim i
  For i = 1 To 100
    AssertEqual "func2:a,b", FuncallObjectMethod(obj_foo, "func2", Array("a", "b"))
  Next
  AssertEqual 100, obj_foo.func2_countCall
End Sub

Sub TestGetObjectMethodSubProc0
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "method0", 0)
  proc
  AssertEqual 1, obj_foo.method0_countCall
End Sub

Sub TestGetObjectMethodSubProc0_manyCall
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "method0", 0)
  Dim i
  For i = 1 To 100
    proc
  Next
  AssertEqual 100, obj_foo.method0_countCall
End Sub

Sub TestGetObjectMethodSubProc0_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodSubProc(obj_foo, "method0", 0)
    proc
  Next
  AssertEqual 100, obj_foo.method0_countCall
End Sub

Sub TestGetObjectMethodSubProc1
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "method1", 1)
  proc "a"
  AssertEqual 1, obj_foo.method1_countCall
End Sub

Sub TestGetObjectMethodSubProc1_manyCall
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "method1", 1)
  Dim i
  For i = 1 To 100
    proc "a"
  Next
  AssertEqual 100, obj_foo.method1_countCall
End Sub

Sub TestGetObjectMethodSubProc1_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodSubProc(obj_foo, "method1", 1)
    proc "a"
  Next
  AssertEqual 100, obj_foo.method1_countCall
End Sub

Sub TestGetObjectMethodSubProc2
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "method2", 2)
  proc "a", "b"
  AssertEqual 1, obj_foo.method2_countCall
End Sub

Sub TestGetObjectMethodSubProc2_manyCall
  Dim proc
  Set proc = GetObjectMethodSubProc(obj_foo, "method2", 2)
  Dim i
  For i = 1 To 100
    proc "a", "b"
  Next
  AssertEqual 100, obj_foo.method2_countCall
End Sub

Sub TestGetObjectMethodSubProc2_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodSubProc(obj_foo, "method2", 2)
    proc "a", "b"
  Next
  AssertEqual 100, obj_foo.method2_countCall
End Sub

Sub TestGetObjectMethodFuncProc0
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "func0", 0)
  AssertEqual "func0", proc
  AssertEqual 1, obj_foo.func0_countCall
End Sub

Sub TestGetObjectMethodFuncProc0_manyCall
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "func0", 0)
  Dim i
  For i = 1 To 100
    AssertEqual "func0", proc
  Next
  AssertEqual 100, obj_foo.func0_countCall
End Sub

Sub TestGetObjectMethodFuncProc0_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodFuncProc(obj_foo, "func0", 0)
    AssertEqual "func0", proc
  Next
  AssertEqual 100, obj_foo.func0_countCall
End Sub

Sub TestGetObjectMethodFuncProc1
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "func1", 1)
  AssertEqual "func1:a", proc("a")
  AssertEqual 1, obj_foo.func1_countCall
End Sub

Sub TestGetObjectMethodFuncProc1_manyCall
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "func1", 1)
  Dim i
  For i = 1 To 100
    AssertEqual "func1:a", proc("a")
  Next
  AssertEqual 100, obj_foo.func1_countCall
End Sub

Sub TestGetObjectMethodFuncProc1_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodFuncProc(obj_foo, "func1", 1)
    AssertEqual "func1:a", proc("a")
  Next
  AssertEqual 100, obj_foo.func1_countCall
End Sub

Sub TestGetObjectMethodFuncProc2
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "func2", 2)
  AssertEqual "func2:a,b", proc("a", "b")
  AssertEqual 1, obj_foo.func2_countCall
End Sub

Sub TestGetObjectMethodFuncProc2_manyCall
  Dim proc
  Set proc = GetObjectMethodFuncProc(obj_foo, "func2", 2)
  Dim i
  For i = 1 To 100
    AssertEqual "func2:a,b", proc("a", "b")
  Next
  AssertEqual 100, obj_foo.func2_countCall
End Sub

Sub TestGetObjectMethodFuncProc2_manyGet
  Dim i, proc
  For i = 1 To 100
    Set proc = GetObjectMethodFuncProc(obj_foo, "func2", 2)
    AssertEqual "func2:a,b", proc("a", "b")
  Next
  AssertEqual 100, obj_foo.func2_countCall
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
