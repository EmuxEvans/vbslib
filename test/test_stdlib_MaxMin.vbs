' stdlib.vbs: Max and Min test.
' @import ../lib/stdlib.vbs

Sub TestNumberCompare_LessThan
  Assert NumberCompare(1, 2) < 0
End Sub

Sub TestNumberCompare_Equal
  Assert NumberCompare(1, 1) = 0
End Sub

Sub TestNumberCompare_GreaterThan
  Assert NumberCompare(2, 1) > 0
End Sub

Class Foo
  Private ivar_name

  Public Property Get Name
    Name = ivar_name
  End Property

  Public Property Let Name(value)
    ivar_name = value
  End Property
End Class

Sub TestObjectPropertyCompare_LessThan
  Dim a: Set a = New Foo: a.Name = "Apple"
  Dim b: Set b = New Foo: b.Name = "Banana"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("TextStringCompare"))

  Assert comp(a, b) < 0
End Sub

Sub TestObjectPropertyCompare_Equal
  Dim a: Set a = New Foo: a.Name = "Apple"
  Dim b: Set b = New Foo: b.Name = "Apple"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("TextStringCompare"))

  Assert comp(a, b) = 0
End Sub

Sub TestObjectPropertyCompare_GreaterThan
  Dim a: Set a = New Foo: a.Name = "Banana"
  Dim b: Set b = New Foo: b.Name = "Apple"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("TextStringCompare"))

  Assert comp(a, b) > 0
End Sub

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
