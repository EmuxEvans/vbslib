' stdlib.vbs: List tool test.
' @import ../lib/stdlib.vbs

Sub TestFind_Found
  AssertEqual "bar", Find(Array("foo", "bar", "baz"), ValueEqual("bar"))
End Sub

Sub TestFind_NotFound
  Assert IsEmpty(Find(Array("foo", "bar"), ValueEqual("baz")))
End Sub

Sub TestFind_EmptyList
  Assert IsEmpty(Find(Array(), ValueEqual("baz")))
End Sub

Sub TestFindPos_Found
  AssertEqual 1, FindPos(Array("foo", "bar", "baz"), ValueEqual("bar"))
End Sub

Sub TestFindPos_NotFound
  Assert IsEmpty(FindPos(Array("foo", "bar"), ValueEqual("baz")))
End Sub

Sub TestFindPos_EmptyList
  Assert IsEmpty(FindPos(Array(), ValueEqual("foo")))
End Sub

Sub TestFindAll_Found
  Dim r
  r = FindAll(Array("foo", "Foo", "FOO", "bar", "baz"), ValueMatch(re("foo", "i")))
  AssertEqual 2, UBound(r)
  AssertEqual "foo", r(0)
  AssertEqual "Foo", r(1)
  AssertEqual "FOO", r(2)
End Sub

Sub TestFindAll_NotFound
  Dim r
  r = FindAll(Array("bar", "baz"), ValueMatch(re("foo", "i")))
  AssertEqual -1, UBound(r)
End Sub

Sub TestFindAll_EmptyList
  Dim r
  r = FindAll(Array(), ValueMatch(re("foo", "i")))
  AssertEqual -1, UBound(r)
End Sub

Sub TestMax
  AssertEqual 7, Max(Array(1, 3, 2, 7, 1, 4), GetRef("NumericCompare"))
End Sub

Sub TestMax_1
  AssertEqual 1, Max(Array(1), GetRef("NumericCompare"))
End Sub

Sub TestMax_Empty
  Assert IsEmpty(Max(Array(), GetRef("NumericCompare")))
End Sub

Sub TestMin
  AssertEqual 1, Min(Array(1, 3, 2, 7, 1, 4), GetRef("NumericCompare"))
End Sub

Sub TestMin_1
  AssertEqual 7, Min(Array(7), GetRef("NumericCompare"))
End Sub

Sub TestMin_Empty
  Assert IsEmpty(Min(Array(), GetRef("NumericCompare")))
End Sub

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
