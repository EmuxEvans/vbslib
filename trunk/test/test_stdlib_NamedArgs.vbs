' stdlib.vbs: Named Arguments test.
' @import ../lib/stdlib.vbs

Option Explicit

Sub TestGetNamedArgumentString_ExistsString
  AssertEqual "Apple", GetNamedArgumentString("foo", D(Array("foo", "Apple")), "Banana")
End Sub

Sub TestGetNamedArgumentString_ExistsEmpty
  Dim optValue, errNum, errSrc
  On Error Resume Next
  optValue = GetNamedArgumentString("foo", D(Array("foo", Empty)), "Banana")
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual RuntimeError, errNum
  AssertEqual "stdlib.vbs:GetNamedArgumentString", errSrc
End Sub

Sub TestGetNamedArgumentString_ExistsBool
  Dim optValue, errNum, errSrc
  On Error Resume Next
  optValue = GetNamedArgumentString("foo", D(Array("foo", True)), "Banana")
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual RuntimeError, errNum
  AssertEqual "stdlib.vbs:GetNamedArgumentString", errSrc
End Sub

Sub TestGetNamedArgumentString_Default
  AssertEqual "Banana", Getnamedargumentstring("foo", D(Array()), "Banana")
End Sub

Sub TestGetNamedArgumentString_NoDefault
  Dim optValue, errNum, errSrc
  On Error Resume Next
  optValue = GetNamedArgumentString("foo", D(Array()), Empty)
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual RuntimeError, errNum
  AssertEqual "stdlib.vbs:GetNamedArgumentString", errSrc
End Sub

Sub TestGetNamedArgumentBool_ExistsTrue
  AssertEqual True, GetNamedArgumentBool("foo", D(Array("foo", True)), Empty)
End Sub

Sub TestGetNamedArgumentBool_ExistsFalse
  AssertEqual False, GetNamedArgumentBool("foo", D(Array("foo", False)), Empty)
End Sub

Sub TestGetNamedArgumentBool_ExistsEmpty
  AssertEqual True, GetNamedArgumentBool("foo", D(Array("foo", Empty)), Empty)
End Sub

Sub TestGetNamedArgumentBool_ExistsString
  Dim optValue, errNum, errSrc
  On Error Resume Next
  optValue = GetNamedArgumentBool("foo", D(Array("foo", "Apple")), Empty)
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual RuntimeError, errNum
  AssertEqual "stdlib.vbs:GetNamedArgumentBool", errSrc
End Sub

Sub TestGetNamedArgumentBool_DefaultTrue
  AssertEqual True, GetNamedArgumentBool("foo", D(Array()), True)
End Sub

Sub TestGetNamedArgumentBool_DefaultFalse
  AssertEqual False, GetNamedArgumentBool("foo", D(Array()), False)
End Sub

Sub TestGetNamedArgumentBool_NoDefault
  Dim optValue, errNum, errSrc
  On Error Resume Next
  optValue = GetNamedArgumentBool("foo", D(Array()), Empty)
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual RuntimeError, errNum
  AssertEqual "stdlib.vbs:GetNamedArgumentBool", errSrc
End Sub

Sub TestGetNamedArgumentSimple_Exists
  AssertEqual True, GetNamedArgumentSimple("foo", D(Array("foo", Empty)))
End Sub

Sub TestGetNamedArgumentSimple_NotExists
  AssertEqual False, GetNamedArgumentSimple("foo", D(Array()))
End Sub

Sub TestGetNamedArgumentSimple_ExistsValue
  Dim optValue, errNum, errSrc
  On Error Resume Next
  optValue = GetNamedArgumentSimple("foo", D(Array("foo", "Apple")))
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual RuntimeError, errNum
  AssertEqual "stdlib.vbs:GetNamedArgumentSimple", errSrc
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
