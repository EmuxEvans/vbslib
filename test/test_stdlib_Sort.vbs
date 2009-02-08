' stdlib.vbs: Sort procedure test.
' @import ../lib/stdlib.vbs

Option Explicit

Sub TestSwapArrayItem_ValueAndValue
  Dim a(1)
  a(0) = "foo"
  a(1) = "bar"
  SwapArrayItem a, 0, 1
  AssertEqual "bar", a(0)
  AssertEqual "foo", a(1)
End Sub

Sub TestSwapArrayItem_ObjectAndObject
  Dim a(1)
  Dim obj1: Set obj1 = CreateObject("Scripting.Dictionary")
  Dim obj2: Set obj2 = CreateObject("Scripting.Dictionary")
  Set a(0) = obj1
  Set a(1) = obj2
  SwapArrayItem a, 0, 1
  AssertSame obj2, a(0)
  AssertSame obj1, a(1)
End Sub

Sub TestNumberCompare_Equal
  Assert NumberCompare(1, 1) = 0
End Sub

Sub TestNumberCompare_Greater
  Assert NumberCompare(2, 1) > 0
End Sub

Sub TestNumberCompare_Less
  Assert NumberCompare(1, 2) < 0
End Sub

Sub TestSort_1
  Dim a(0)
  a(0) = 0
  Sort a, NumberCompare
  AssertEqual 0, a(0)
End Sub

Sub TestSort_2
  Dim a(1)
  a(0) = 1
  a(1) = 0
  Sort a, NumberCompare
  AssertEqual 0, a(0)
  AssertEqual 1, a(1)
End Sub

Sub TestSort_3
  Dim a(2)
  a(0) = 1
  a(1) = 3
  a(2) = 0
  Sort a, NumberCompare
  AssertEqual 0, a(0)
  AssertEqual 1, a(1)
  AssertEqual 3, a(2)
End Sub

Const COUNT_OF_MANY_ITEMS = 2000

Sub TestSort_Many
  Dim i
  ReDim a(COUNT_OF_MANY_ITEMS - 1)

  For i = 0 To COUNT_OF_MANY_ITEMS - 1
    a(i) = Int(Rnd * COUNT_OF_MANY_ITEMS)
  Next

  Sort a, NumberCompare

  For i = 0 To COUNT_OF_MANY_ITEMS - 2
    Assert a(i) <= a(i + 1)
  Next
End Sub

Sub TestSort_SortedList
  Dim i
  ReDim a(COUNT_OF_MANY_ITEMS - 1)

  For i = 0 To COUNT_OF_MANY_ITEMS - 1
    a(i) = Int(Rnd * COUNT_OF_MANY_ITEMS)
  Next

  Sort a, NumberCompare
  Sort a, NumberCompare

  For i = 0 To COUNT_OF_MANY_ITEMS - 2
    Assert a(i) <= a(i + 1)
  Next
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
