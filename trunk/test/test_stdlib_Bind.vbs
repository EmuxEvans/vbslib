' stdlib.vbs: Bind subroutine test.
' @import ../lib/stdlib.vbs

Option Explicit

Sub TestBind_Value
  Dim v
  Bind v, "Hello world."
  AssertEqual "Hello world.", v
End Sub

Sub TestBind_Object
  Dim v
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  Bind v, obj
  AssertSame obj, v
End Sub

Sub TestBindAt_Value
  Dim a(1)
  BindAt a, 0, "Hello world."
  AssertEqual "Hello world.", a(0)
  Assert IsEmpty(a(1))
End Sub

Sub TestBindAt_Object
  Dim a(1)
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  BindAt a, 0, Nothing
  AssertSame Nothing, a(0)
  Assert IsEmpty(a(1))
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
