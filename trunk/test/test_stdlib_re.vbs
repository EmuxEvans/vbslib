' stdlib.vbs: RegExp utility test.
' @import ../lib/stdlib.vbs

Sub TestRe_NoOption
  Dim regex
  Set regex = re("foo", "")
  AssertEqual True, regex.Test("foo")
  AssertEqual False, regex.Test("bar")
End Sub

Sub TestRe_IgnoreCase
  Dim regex
  Set regex = re("foo", "i")
  AssertEqual True, regex.Test("foo")
  AssertEqual True, regex.Test("Foo")
  AssertEqual True, regex.Test("FOO")
End Sub

Sub TestRe_NoIgnoreCase
  Dim regex
  Set regex = re("foo", "")
  AssertEqual True, regex.Test("foo")
  AssertEqual False, regex.Test("Foo")
  AssertEqual False, regex.Test("FOO")
End Sub

Sub TestRe_Global
  Dim regex
  Set regex = re("foo", "g")
  AssertEqual 1, regex.Execute("foo").Count
  AssertEqual 2, regex.Execute("foo,foo").Count
  AssertEqual 3, regex.Execute("foo,foo,foo").Count
End Sub

Sub TestRe_NoGlobal
  Dim regex
  Set regex = re("foo", "")
  AssertEqual 1, regex.Execute("foo").Count
  AssertEqual 1, regex.Execute("foo,foo").Count
  AssertEqual 1, regex.Execute("foo,foo,foo").Count
End Sub

Sub TestRe_Multiline
  Dim regex
  Set regex = re("^foo$", "gm")
  AssertEqual 1, regex.Execute(Join(Array("foo"), vbNewLine)).Count
  AssertEqual 2, regex.Execute(Join(Array("foo", "foo"), vbNewLine)).Count
  AssertEqual 3, regex.Execute(Join(Array("foo", "foo", "foo"), vbNewLine)).Count
End Sub

Sub TestRe_NoMultiline
  Dim regex
  Set regex = re("^foo$", "g")
  AssertEqual 1, regex.Execute(Join(Array("foo"), vbNewLine)).Count
  AssertEqual 0, regex.Execute(Join(Array("foo", "foo"), vbNewLine)).Count
  AssertEqual 0, regex.Execute(Join(Array("foo", "foo", "foo"), vbNewLine)).Count
End Sub

Sub TestRe_AllOptions
  Dim regex
  Set regex = re("^foo$", "igm")
  AssertEqual 3, regex.Execute(Join(Array("foo", "Foo", "FOO"), vbNewLine)).count
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
