' stdlib.vbs: list search test.
' @import ../lib/stdlib.vbs

Sub TestRegExpReplace
  Dim func
  Set func = RegExpReplace(re("xyz", ""), "XYZ")
  AssertEqual "abcXYZdef", func("abcxyzdef")
End Sub

Sub TestMap_RegExpReplace
  Dim a
  a = Map(Array("foo", "bar", "baz"), RegExpReplace(re("ba", ""), "BA"))
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
