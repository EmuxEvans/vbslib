' stdlib.vbs: Utility Function test.
' @import ../lib/stdlib.vbs

Option Explicit

Class Foo
  Private ivar_name

  Public Property Get Name
    Name = ivar_name
  End Property

  Public Property Let Name(value)
    ivar_name = value
  End Property
End Class

Class Bar
  Private ivar_foo
  Private ivar_bar
  Private ivar_baz

  Public Property Get Foo
    Foo = ivar_foo
  End Property

  Public Property Let Foo(value)
    ivar_foo = value
  End Property

  Public Property Get Bar
    Bar = ivar_bar
  End Property

  Public Property Let Bar(value)
    ivar_bar = value
  End Property

  Public Property Get Baz
    Baz = ivar_baz
  End Property

  Public Property Let Baz(value)
    ivar_baz = value
  End Property
End Class

Sub TestValueReplace
  Dim func
  Set func = ValueReplace(re("xyz", ""), "XYZ")
  AssertEqual "abcXYZdef", func("abcxyzdef")
End Sub

Sub TestValueItemAt
  Dim func
  Set func = ValueItemAt("bar")
  AssertEqual "Banana", func(D(Array("foo", "Apple", "bar", "Banana", "baz", "Orange")))
End Sub

Sub TestValueItemsAt
  Dim func
  Set func = ValueItemsAt(Array("foo", "baz"))
  AssertEqual ShowValue(D(Array("foo", "Apple", "baz", "Orange"))), _
              ShowValue(func(D(Array("foo", "Apple", "bar", "Banana", "baz", "Orange"))))
End Sub

Sub TestValueObjectProperty
  Dim func
  Set func = ValueObjectProperty("Name")

  Dim foo
  Set foo = New Foo
  foo.Name = "bar"

  AssertEqual "bar", func(foo)
End Sub

Sub TestValueObjectProperties
  Dim func
  Set func = ValueObjectProperties(Array("Foo", "Baz"))

  Dim o
  Set o = New Bar
  o.Foo = "Apple"
  o.Bar = "Banana"
  o.Baz = "Orange"

  AssertEqual ShowValue(D(Array("Foo", "Apple", "Baz", "Orange"))), ShowValue(func(o))
End Sub

Sub TestPriorMax
  AssertEqual "Apple", PriorMax(GetRef("StrComp_"), "Apple", "Apple")
  AssertEqual "Banana", PriorMax(GetRef("StrComp_"), "Apple", "Banana")
  AssertEqual "Banana", PriorMax(GetRef("StrComp_"), "Banana", "Apple")
End Sub

Sub TestPriorMin
  AssertEqual "Apple", PriorMin(GetRef("StrComp_"), "Apple", "Apple")
  AssertEqual "Apple", PriorMin(GetRef("StrComp_"), "Apple", "Banana")
  AssertEqual "Apple", PriorMin(GetRef("StrComp_"), "Banana", "Apple")
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

Sub TestReverseCompare_GreaterThan
  Dim comp
  Set comp = ReverseCompare(GetRef("StrComp_"))
  Assert StrComp("Apple", "Banana") < 0
  Assert comp("Apple", "Banana") > 0
End Sub

Sub TestReverseCompare_Equal
  Dim comp
  Set comp = ReverseCompare(GetRef("StrComp_"))
  Assert StrComp("Apple", "Apple") = 0
  Assert comp("Apple", "Apple") = 0
End Sub

Sub TestReverseCompare_LessThan
  Dim comp
  Set comp = ReverseCompare(GetRef("StrComp_"))
  Assert StrComp("Banana", "Apple") > 0
  Assert comp("Banana", "Apple") < 0
End sub

Sub TestCompareFilter_Equal
  Dim comp
  Set comp = CompareFilter(GetRef("CInt_"), GetRef("NumericCompare"))
  Assert "010" <> "10"
  Assert comp("010", "10") = 0
End Sub

Sub TestCompareFilter_GreaterThan
  Dim comp
  Set comp = CompareFilter(GetRef("CInt_"), GetRef("NumericCompare"))
  Assert "010" < "1"
  Assert comp("010", "1") > 0
End Sub

Sub TestCompareFilter_LessThan
  Dim comp
  Set comp = CompareFilter(GetRef("CInt_"), GetRef("NumericCompare"))
  Assert "1" > "010"
  Assert comp("1", "010") < 0
End Sub

Sub TestObjectPropertyCompare_Equal
  Dim a: Set a = New Foo: a.Name = "Apple"
  Dim b: Set b = New Foo: b.Name = "Apple"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("StrComp_"))

  Assert comp(a, b) = 0
End Sub

Sub TestObjectPropertyCompare_LessThan
  Dim a: Set a = New Foo: a.Name = "Apple"
  Dim b: Set b = New Foo: b.Name = "Banana"

  Dim comp
  Set comp = ObjectPropertyCompare("Name", GetRef("StrComp_"))

  Assert comp(a, b) < 0
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
  Set gt = ValueGreaterThan(0)
  Assert gt(1)
  Assert Not gt(0)
  Assert Not gt(-1)
End Sub

Sub TestValueGreaterEqual
  Dim ge
  Set ge = ValueGreaterEqual(0)
  Assert ge(1)
  Assert ge(0)
  Assert Not ge(-1)
End Sub

Sub TestValueLessThan
  Dim lt
  Set lt = ValueLessThan(0)
  Assert lt(-1)
  Assert Not lt(0)
  Assert Not lt(1)
End Sub

Sub TestValueLessEqual
  Dim le
  Set le = ValueLessEqual(0)
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
  Set comp = ValueCompare("=", "Banana", GetRef("StrComp_"))

  Assert Not comp("Apple")
  Assert comp("Banana")
  Assert Not comp("Orange")
End Sub

Sub TestValueCompare_GreaterThan
  Dim comp
  Set comp = ValueCompare(">", "Banana", GetRef("StrComp_"))

  Assert Not comp("Apple")
  Assert Not comp("Banana")
  Assert comp("Orange")
End Sub

Sub TestValueCompare_GreaterEqual
  Dim comp
  Set comp = ValueCompare(">=", "Banana", GetRef("StrComp_"))

  Assert Not comp("Apple")
  Assert comp("Banana")
  Assert comp("Orange")
End Sub

Sub TestValueCompare_LessThan
  Dim comp
  Set comp = ValueCompare("<", "Banana", GetRef("StrComp_"))

  Assert comp("Apple")
  Assert Not comp("Banana")
  Assert Not comp("Orange")
End Sub

Sub TestValueCompare_LessEqual
  Dim comp
  Set comp = ValueCompare("<=", "Banana", GetRef("StrComp_"))

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

Sub TestAdd
  AssertEqual 3, Add(1, 2)
  AssertEqual -1, Add(1, -2)
End Sub

Sub TestSubtract
  AssertEqual 1, Subtract(2, 1)
  AssertEqual -1, Subtract(1, 2)
End Sub

Sub TestMultiply
  AssertEqual 6, Multiply(2, 3)
  AssertEqual -6, Multiply(2, -3)
  AssertEqual 6, Multiply(-2, -3)
End Sub

Sub TestDivide
  AssertEqual 3, Divide(6, 2)
  AssertEqual 1.5, Divide(3, 2)
End Sub

Sub TestMod_
  AssertEqual 2, Mod_(5, 3)
End Sub

Sub TestPower
  AssertEqual 1, Power(2, 0)
  AssertEqual 2, Power(2, 1)
  AssertEqual 4, Power(2, 2)
  AssertEqual 8, Power(2, 3)
  AssertEqual 16, Power(2, 4)

  AssertEqual 1, Power(-2, 0)
  AssertEqual -2, Power(-2, 1)
  AssertEqual 4, Power(-2, 2)
  AssertEqual -8, Power(-2, 3)
  AssertEqual 16, Power(-2, 4)
End Sub

Sub TestConcat
  AssertEqual "foo,bar", Concat("foo", ",bar")
End Sub

Sub TestNot_
  AssertEqual -1, Not_(0)
End Sub

Sub TestAnd_
  AssertEqual 2, And_(1 + 2 + 4, 2)
End Sub

Sub TestOr_
  AssertEqual 1 + 2 + 4, Or_(2, 1 + 4)
End Sub

Sub TestXor_
  AssertEqual 2 + 4, Xor_(1 + 4, 1 + 2)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
