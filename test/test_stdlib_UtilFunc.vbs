' stdlib.vbs: Utility Function test.
' @import ../lib/stdlib.vbs

Class Foo
  Private ivar_name

  Public Property Get Name
    Name = ivar_name
  End Property

  Public Property Let Name(value)
    ivar_name = value
  End Property
End Class

Sub TestValueReplace
  Dim func
  Set func = ValueReplace(re("xyz", ""), "XYZ")
  AssertEqual "abcXYZdef", func("abcxyzdef")
End Sub

Sub TestValueObjectProperty
  Dim func
  Set func = ValueObjectProperty("Name")

  Dim foo
  Set foo = New Foo
  foo.Name = "bar"

  AssertEqual "bar", func(foo)
End Sub

Sub TestValueDictionaryItem
  Dim func
  Set func = ValueDictionaryItem("bar")

  Dim d
  Set d = CreateObject("Scripting.Dictionary")
  d("foo") = "Apple"
  d("bar") = "Banana"
  d("baz") = "Orange"

  AssertEqual "Banana", func(d)
End Sub

Sub TestNumberCompare_LessThan
  Assert NumberCompare(1, 2) < 0
End Sub

Sub TestNumberCompare_Equal
  Assert NumberCompare(1, 1) = 0
End Sub

Sub TestNumberCompare_GreaterThan
  Assert NumberCompare(2, 1) > 0
End Sub

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

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
