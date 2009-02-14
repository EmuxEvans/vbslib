' stdlib.vbs: Find test.
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

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
