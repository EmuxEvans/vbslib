' stdlib.vbs: ListBuffer class test.
' @import ../lib/stdlib.vbs

Option Explicit

Dim listBuf

Sub SetUp
  Set listBuf = New ListBuffer
End Sub

Sub TearDown
  Set listBuf = Nothing
End Sub

Sub TestListBuffer_Count
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  AssertEqual 3, listBuf.Count
End Sub

Sub TestListBuffer_CountEmpty
  AssertEqual 0, listBuf.Count
End Sub

Sub TestListBuffer_GetItem
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  AssertEqual "foo", listBuf(0)
  AssertEqual "bar", listBuf(1)
  AssertSame obj, listBuf(2)
End Sub

Sub TestListBuffer_GetItemOutOfRange
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  Dim errNum, errSrc
  On Error Resume Next
  listBuf(3)
  errNum = Err.Number
  errSrc = Err.source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:ListBuffer.Item(Get)", errSrc
End Sub

Sub TestListBuffer_GetItemEmpty
  Dim errNum, errSrc
  On Error Resume Next
  listBuf(0)
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:ListBuffer.Item(Get)", errSrc
End Sub

Sub TestListBuffer_LetItem
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  listBuf(2) = "baz"
  AssertEqual "foo", listBuf(0)
  AssertEqual "bar", listBuf(1)
  AssertEqual "baz", listBuf(2)
End Sub

Sub TestListBuffer_LetItemOutOfRange
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  Dim errNum, errSrc
  On Error Resume Next
  listBuf(3) = "baz"
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:ListBuffer.Item(Let)", errSrc
End Sub

Sub TestListBuffer_LetItemEmpty
  Dim errNum, errSrc
  On Error Resume Next
  listBuf(0) = "foo"
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:ListBuffer.Item(Let)", errSrc
End Sub

Sub TestListBuffer_SetItem
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  Set listBuf(0) = obj
  AssertSame obj, listBuf(0)
  AssertEqual "bar", listBuf(1)
  AssertSame obj, listBuf(2)
End Sub

Sub TestListBuffer_SetItemOutOfRange
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  Dim errNum, errSrc
  On Error Resume Next
  Set listBuf(3) = Nothing
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:ListBuffer.Item(Set)", errSrc
End Sub

Sub TestListBuffer_SetItemEmpty
  Dim errNum, errSrc
  On Error Resume Next
  Set listBuf(0) = Nothing
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 9, errNum
  AssertEqual "stdlib.vbs:ListBuffer.Item(Set)", errSrc
End Sub

Sub TestListBuffer_Items
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  Dim a
  a = listBuf.Items
  AssertEqual "Variant()", TypeName(a)
  AssertEqual 0, LBound(a)
  AssertEqual 2, UBound(a)
  AssertEqual "foo", a(0)
  AssertEqual "bar", a(1)
  AssertSame obj, a(2)
End Sub

Sub TestListBuffer_ItemsEmpty
  Dim a
  a = listBuf.Items
  AssertEqual "Variant()", TypeName(a)
  AssertEqual 0, LBound(a)
  AssertEqual -1, UBound(a)
End Sub

Sub TestListBuffer_Exists
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  Assert listBuf.Exists(0)
  Assert listBuf.Exists(1)
  Assert listBuf.Exists(2)
  Assert Not listBuf.Exists(3)
End Sub

Sub TestListBuffer_ExistsEmpty
  Assert Not listBuf.Exists(0)
End Sub

Sub TestListBuffer_Append
  listBuf.Append Array(1, 2, 3, "a", "b")
  AssertEqual 5, listBuf.Count
  AssertEqual 1, listBuf(0)
  AssertEqual 2, listBuf(1)
  AssertEqual 3, listBuf(2)
  AssertEqual "a", listBuf(3)
  AssertEqual "b", listBuf(4)
End Sub

Sub TestListBuffer_RemoveAll
  Dim obj: Set obj = CreateObject("Scripting.Dictionary")
  listBuf.Add "foo"
  listBuf.Add "bar"
  listBuf.Add obj

  listBuf.RemoveAll
  AssertEqual 0, listBuf.Count
  Assert Not listBuf.Exists(0)
  Assert Not listBuf.Exists(1)
  Assert Not listBuf.Exists(2)
End Sub

Sub TestListBuffer_RemoveAllEmpty
  listBuf.RemoveAll
  AssertEqual 0, listBuf.Count
  Assert Not listBuf.Exists(0)
  Assert Not listBuf.Exists(1)
  Assert Not listBuf.Exists(2)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
