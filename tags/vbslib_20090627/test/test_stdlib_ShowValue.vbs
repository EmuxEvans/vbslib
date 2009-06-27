' stdlib.vbs: ShowValue procedure test.
' @import ../lib/stdlib.vbs

Option Explicit

Class Foo
End Class

Sub TestShowValue_Empty
  AssertEqual "<empty>", ShowValue(Empty)
End Sub

Sub TestShowValue_Null
  AssertEqual "<null>", ShowValue(Null)
End Sub

Sub TestShowValue_BoolTrue
  AssertEqual "True", ShowValue(True)
End Sub

Sub TestShowValue_BoolFalse
  AssertEqual "False", ShowValue(False)
End Sub

Sub TestShowValue_Byte
  AssertEqual "1", ShowValue(CByte(1))
End Sub

Sub TestShowValue_Int
  AssertEqual "1000", ShowValue(CInt(1000))
End Sub

Sub TestShowValue_Currency
  AssertEqual "1000", ShowValue(CCur(1000))
End Sub

Sub TestShowValue_Long
  AssertEqual "100000", ShowValue(CLng(100000))
End Sub

Sub TestShowValue_Single
  AssertEqual "1.23", ShowValue(CSng(1.23))
End Sub

Sub TestShowValue_Double
  AssertEqual "1.23", ShowValue(CDbl(1.23))
End Sub

Sub TestShowValue_Date
  AssertEqual "2009/01/24 18:12:04", ShowValue(#2009-01-24 18:12:04#)
End Sub

Sub TestShowValue_String
  AssertEqual """foo""", ShowValue("foo")
End Sub

Sub TestShowValue_StringQuote
  AssertEqual """foo""""bar""", ShowValue("foo""bar")
End Sub

Sub TestShowValue_Object
  AssertEqual "<Foo>", ShowValue(New Foo)
End Sub

Sub TestShowValue_ObjectDictionary
  Dim d: Set d = CreateObject("Scripting.Dictionary")
  d("foo") = "Apple"
  d("bar") = "Banana"
  d("baz") = "Orange"
  AssertEqual "{""foo""=>""Apple"",""bar""=>""Banana"",""baz""=>""Orange""}", ShowValue(d)
End Sub

Sub TestShowValue_ObjectDictionaryEmpty
  AssertEqual "{}", ShowValue(CreateObject("Scripting.Dictionary"))
End Sub

Sub TestShowValue_Array
  Dim a(5)
  a(0) = 1
  a(1) = "foo"
  ' a(2) is Empty.
  a(3) = Null
  Set a(4) = CreateObject("Scripting.Dictionary")
  a(4)("bar") = "Banana"

  Dim b(1)
  b(0) = "a"
  b(1) = "b"
  a(5) = b

  AssertEqual "[1,""foo"",<empty>,<null>,{""bar""=>""Banana""},[""a"",""b""]]", ShowValue(a)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
