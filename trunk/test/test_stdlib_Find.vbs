' stdlib.vbs: Find test.
' @import ../lib/stdlib.vbs

Sub TestValueEqual
  Dim eq
  Set eq = ValueEqual("foo")
  Assert eq("foo")
  Assert Not eq("bar")
End Sub

Sub TestValueGreaterThan
  Dim gt
  Set gt = ValueGreaterThan(0, True)
  Assert gt(1)
  Assert Not gt(0)
  Assert Not gt(-1)
End Sub

Sub TestValueGreaterThanEqual
  Dim ge
  Set ge = ValueGreaterThan(0, False)
  Assert ge(1)
  Assert ge(0)
  Assert Not ge(-1)
End Sub

Sub TestValueLessThan
  Dim lt
  Set lt = ValueLessThan(0, True)
  Assert lt(-1)
  Assert Not lt(0)
  Assert Not lt(1)
End Sub

Sub TestValueLessThanEqual
  Dim le
  Set le = ValueLessThan(0, False)
  Assert le(-1)
  Assert le(0)
  Assert Not le(1)
End Sub

Sub TestValueBetween
  Dim btwn
  Set btwn = ValueBetween(2, 4, False)
  Assert Not btwn(0)
  Assert Not btwn(1)
  Assert btwn(2)
  Assert btwn(3)
  Assert btwn(4)
  Assert Not btwn(5)
End Sub

Sub TestValueBetweenExcludeUpperBound
  Dim btwn
  Set btwn = ValueBetween(2, 4, True)
  Assert Not btwn(0)
  Assert Not btwn(1)
  Assert btwn(2)
  Assert btwn(3)
  Assert Not btwn(4)
  Assert Not btwn(5)
End Sub

Sub TestValueMatch
  Dim ma
  Set ma = ValueMatch(re("foo", ""))
  Assert ma("foo, bar, baz")
  Assert Not ma("bar, baz")
End Sub

Sub TestValueFilter
  Dim f
  Set f = ValueFilter(GetRef("UCase_"), ValueEqual("FOO"))
  Assert f("foo")
  Assert f("Foo")
  Assert f("FOO")
  Assert Not f("bar")
End Sub

Sub TestNotCond
  Dim cond
  Set cond = NotCond(ValueEqual("foo"))
  Assert Not cond("foo")
  Assert cond("bar")
End Sub

Sub TestAndCond
  Dim cond
  Set cond = AndCond(ValueMatch(re("foo", "")), ValueMatch(re("bar", "")))
  Assert Not cond("foo")
  Assert Not cond("bar")
  Assert cond("foo, bar")
End Sub

Sub TestOrCond
  Dim cond
  Set cond = OrCond(ValueMatch(re("foo", "")), ValueMatch(re("bar", "")))
  Assert cond("foo")
  Assert cond("bar")
  Assert Not cond("baz")
End Sub

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
