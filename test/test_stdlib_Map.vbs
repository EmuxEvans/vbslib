' stdlib.vbs: Map test.
' @import ../lib/stdlib.vbs

Sub TestMap_ValueReplace
  Dim a
  a = Map(Array("foo", "bar", "baz"), ValueReplace(re("ba", ""), "BA"))
  AssertEqual 2, UBound(a)
  AssertEqual "foo", a(0)
  AssertEqual "BAr", a(1)
  AssertEqual "BAz", a(2)
End Sub

Sub TestMap_VBScriptFunctionAlias
  Dim a
  a = Map(Array(" foo", "bar ", " baz "), GetRef("Trim_"))
  AssertEqual 2, UBound(a)
  AssertEqual "foo", a(0)
  AssertEqual "bar", a(1)
  AssertEqual "baz", a(2)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
