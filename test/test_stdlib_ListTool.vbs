' stdlib.vbs: List tool test.
' @import ../lib/stdlib.vbs

Option Explicit

Sub TestFirstItem
  AssertEqual "foo", FirstItem(Array("foo", "bar", "baz"))
End Sub

Sub TestFirstItem_EmptyList
  Dim errNum, errSrc
  On Error Resume Next
  FirstItem(Array())
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:FirstItem", errSrc
End Sub

Sub TestLastItem
  AssertEqual "baz", LastItem(Array("foo", "bar", "baz"))
End Sub

Sub TestLastItem_EmptyList
  Dim errNum, errSrc
  On Error Resume Next
  LastItem(Array())
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:LastItem", errSrc
End Sub

Sub TestCountItem_Array
  AssertEqual 3, CountItem(Array("foo", "bar", "baz"))
End Sub

Sub TestCountItem_Object
  Dim d
  Set d = CreateObject("Scripting.Dictionary")
  d("foo") = "Apple"
  d("bar") = "Banana"
  d("baz") = "Orange"

  AssertEqual 3, CountItem(d)
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

Sub TestRange
  Dim a
  a = Range(0, ValueLessThan(10), GetFuncProcSubset(GetRef("Add"), 2, Array(1)))

  AssertEqual 9, UBound(a)
  AssertEqual 0, a(0)
  AssertEqual 1, a(1)
  AssertEqual 2, a(2)
  AssertEqual 3, a(3)
  AssertEqual 4, a(4)
  AssertEqual 5, a(5)
  AssertEqual 6, a(6)
  AssertEqual 7, a(7)
  AssertEqual 8, a(8)
  AssertEqual 9, a(9)
End Sub

Sub TestRange_Empty
  Dim a
  a = Range(0, ValueLessThan(0), GetFuncProcSubset(GetRef("Add"), 2, Array(1)))
  AssertEqual -1, UBound(a)
End Sub

Sub TestNumericRange
  Dim a
  a = NumericRange(0, 10, 2, False)

  AssertEqual 5, UBound(a)
  AssertEqual 0, a(0)
  AssertEqual 2, a(1)
  AssertEqual 4, a(2)
  AssertEqual 6, a(3)
  AssertEqual 8, a(4)
  AssertEqual 10, a(5)
End Sub

Sub TestNumericRange_Exclude
  Dim a
  a = NumericRange(0, 10, 2, True)

  AssertEqual 4, UBound(a)
  AssertEqual 0, a(0)
  AssertEqual 2, a(1)
  AssertEqual 4, a(2)
  AssertEqual 6, a(3)
  AssertEqual 8, a(4)
End Sub

Sub TestNumericRange_Empty
  Dim a
  a = NumericRange(0, 0, 2, True)
  AssertEqual -1, UBound(a)
End Sub

Sub TestNumbering
  Dim a
  a = Numbering(1, 5)

  AssertEqual 4, UBound(a)
  AssertEqual 1, a(0)
  AssertEqual 2, a(1)
  AssertEqual 3, a(2)
  AssertEqual 4, a(3)
  AssertEqual 5, a(4)
End Sub

Sub TestNumbering_Empty
  Dim a
  a = Numbering(1, 0)
  AssertEqual -1, UBound(a)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
