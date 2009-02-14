' stdlib.vbs: Object Accessor test.
' @import ../lib/stdlib.vbs

Option Explicit

Class Foo
  Private ivar_bar

  Public Property Get Bar
    Bind Bar, ivar_bar
  End Property

  Public Property Let Bar(value)
    ivar_bar = value
  End Property

  Public Property Set Bar(value)
    Set ivar_bar = value
  End Property
End Class

Class Bar
End Class

Dim obj_foo, obj_bar

Sub SetUp
  Set obj_foo = New Foo
  Set obj_bar = New Bar
End Sub

Sub TearDown
  Set obj_foo = Nothing
  Set obj_bar = Nothing
End Sub

Sub TestGetObjectProeprty_Value
  obj_foo.Bar = "Hello world."
  AssertEqual "Hello world.", GetObjectProperty(obj_foo, "Bar")
End Sub

Sub TestGetObjectProeprty_Object
  Set obj_foo.Bar = obj_bar
  AssertSame obj_bar, GetObjectProperty(obj_foo, "Bar")
End Sub

Sub TestGetObjectProperty_NoProperty
  Dim v, errNum
  On Error Resume Next
  v = GetObjectProperty(obj_bar, "Baz")
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  AssertEqual 438, errNum
End Sub

Sub TestSetObjectProperty_Value
  SetObjectProperty obj_foo, "Bar", "Hello world."
  AssertEqual "Hello world.", obj_foo.Bar
End Sub

Sub TestSetObjectProperty_ValueNoProperty
  Dim errNum
  On Error Resume Next
  SetObjectProperty obj_foo, "Baz", "Hello world."
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  AssertEqual 438, errNum
End Sub

Sub TestSetObjectProperty_Object
  SetObjectProperty obj_foo, "Bar", obj_bar
  AssertSame obj_bar, obj_foo.Bar
End Sub

Sub TestSetObjectProperty_ObjectNoProperty
  Dim errNum
  On Error Resume Next
  SetObjectProperty obj_foo, "Baz", obj_bar
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  AssertEqual 438, errNum
End Sub

Sub TestObjectPropertyExists_Exists
  AssertEqual True, ObjectPropertyExists(obj_foo, "Bar")
End Sub

Sub TestObjectPropertyExists_NotExists
  AssertEqual False, ObjectPropertyExists(obj_foo, "Baz")
End Sub

Sub TestObjectPropertyUpDownCase
  SetObjectProperty obj_foo, "bar", "Hello world."
  AssertEqual "Hello world.", GetObjectProperty(obj_foo, "BAR")
  AssertEqual True, ObjectPropertyExists(obj_foo, "bAr")
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
