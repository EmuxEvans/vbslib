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

Sub TestNumericCompare_LessThan
  Assert NumericCompare(1, 2) < 0
End Sub

Sub TestNumericCompare_Equal
  Assert NumericCompare(1, 1) = 0
End Sub

Sub TestNumericCompare_GreaterThan
  Assert NumericCompare(2, 1) > 0
End Sub

Sub TestObjectPropertyCompare_LessThan
  Dim a: Set a = New Foo: a.Name = "Apple"
  Dim b: Set b = New Foo: b.Name = "Banana"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("StrComp_"))

  Assert comp(a, b) < 0
End Sub

Sub TestObjectPropertyCompare_Equal
  Dim a: Set a = New Foo: a.Name = "Apple"
  Dim b: Set b = New Foo: b.Name = "Apple"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("StrComp_"))

  Assert comp(a, b) = 0
End Sub

Sub TestObjectPropertyCompare_GreaterThan
  Dim a: Set a = New Foo: a.Name = "Banana"
  Dim b: Set b = New Foo: b.Name = "Apple"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("StrComp_"))

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

Sub TestValueCompare_Equal
  Dim comp
  Set comp = ValueCompare("eq", "Banana", GetRef("StrComp_"))

  Assert Not comp("Apple")
  Assert comp("Banana")
  Assert Not comp("Orange")
End Sub

Sub TestValueCompare_GreaterThan
  Dim comp
  Set comp = ValueCompare("gt", "Banana", GetRef("StrComp_"))

  Assert Not comp("Apple")
  Assert Not comp("Banana")
  Assert comp("Orange")
End Sub

Sub TestValueCompare_GreaterThanEqual
  Dim comp
  Set comp = ValueCompare("ge", "Banana", GetRef("StrComp_"))

  Assert Not comp("Apple")
  Assert comp("Banana")
  Assert comp("Orange")
End Sub

Sub TestValueCompare_LessThan
  Dim comp
  Set comp = ValueCompare("lt", "Banana", GetRef("StrComp_"))

  Assert comp("Apple")
  Assert Not comp("Banana")
  Assert Not comp("Orange")
End Sub

Sub TestValueCompare_LessThanEqual
  Dim comp
  Set comp = ValueCompare("le", "Banana", GetRef("StrComp_"))

  Assert comp("Apple")
  Assert comp("Banana")
  Assert Not comp("Orange")
End Sub

Sub TestValueCompare_UnknownOperatorType
  Dim comp, errNum, errSrc
  On Error Resume Next
  Set comp = ValueCompare("foo", "Banana", GetRef("StrComp_"))
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 5, errNum
  AssertEqual "stdlib.vbs:ValueCompare", errSrc
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
  Assert Not cond("baz")
End Sub

Sub TestOrCond
  Dim cond
  Set cond = OrCond(ValueMatch(re("foo", "")), ValueMatch(re("bar", "")))
  Assert cond("foo")
  Assert cond("bar")
  Assert cond("foo, bar")
  Assert Not cond("baz")
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
