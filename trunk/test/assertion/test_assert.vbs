' assertion example
' @import ../../lib/stdlib.vbs

Option Explicit

Class Foo
End Class

Sub TestAssert_OK
  Assert True
End Sub

Sub TestAssert_NG
  Assert False
End Sub

Sub TestAssertEqual_OK
  AssertEqual "foo", "foo"
End Sub

Sub TestAssertEqual_NG
  AssertEqual "foo", "bar"
End Sub

Sub TestAssertSame_OK
  Dim obj
  Set obj = New Foo
  AssertSame obj, obj
End Sub

Sub TestAssertSame_NG
  AssertSame New Foo, New Foo
End Sub

Sub TestAssertMatch_OK
  AssertMatch re("foo", "i"), "foo"
  AssertMatch re("foo", "i"), "FOO"
  AssertMatch re("foo", "i"), "foo, bar, baz"
End Sub

Sub TestAssertMatch_NG
  AssertMatch re("foo", "i"), "bar"
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
