' stdlib.vbs: Map test.
' @import ../lib/stdlib.vbs

Class Foo
  Private ivar_bar

  Public Property Get Bar
    Bar = ivar_bar
  End Property

  Public Property Let Bar(value)
    ivar_bar = value
  End Property
End Class

Sub TestValueReplace
  Dim func
  Set func = ValueReplace(re("xyz", ""), "XYZ")
  AssertEqual "abcXYZdef", func("abcxyzdef")
End Sub

Sub TestValueObjectProperty
  Dim func
  Set func = ValueObjectProperty("Bar")

  Dim foo
  Set foo = New Foo
  foo.Bar = "Hello world."

  AssertEqual "Hello world.", func(foo)
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
