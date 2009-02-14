' stdlib.vbs: Max and Min test.
' @import ../lib/stdlib.vbs

Sub TestMax
  AssertEqual 7, Max(Array(1, 3, 2, 7, 1, 4), GetRef("NumberCompare"))
End Sub

Sub TestMax_1
  AssertEqual 1, Max(Array(1), GetRef("NumberCompare"))
End Sub

Sub TestMax_Empty
  Assert IsEmpty(Max(Array(), GetRef("NumberCompare")))
End Sub

Sub TestMin
  AssertEqual 1, Min(Array(1, 3, 2, 7, 1, 4), GetRef("NumberCompare"))
End Sub

Sub TestMin_1
  AssertEqual 7, Min(Array(7), GetRef("NumberCompare"))
End Sub

Sub TestMin_Empty
  Assert IsEmpty(Min(Array(), GetRef("NumberCompare")))
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
