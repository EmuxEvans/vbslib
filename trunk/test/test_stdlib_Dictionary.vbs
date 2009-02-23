' stdlib.vbs: Dictionary procedure test.
' @import ../lib/stdlib.vbs

Option Explicit

Sub TestDictionary_Even
  Dim dict
  Set dict = Dictionary(Array("foo", "Apple", "bar", "Banana", "baz", "Orange"))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionary_Even_shortcut
  Dim dict
  Set dict = D(Array("foo", "Apple", "bar", "Banana", "baz", "Orange"))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionary_Odd
  Dim dict
  Set dict = Dictionary(Array("foo", "Apple", "bar", "Banana", "baz"))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
  Assert IsEmpty(dict("baz"))
End Sub

Sub TestDictionary_Odd_shortcut
  Dim dict
  Set dict = D(Array("foo", "Apple", "bar", "Banana", "baz"))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
  Assert IsEmpty(dict("baz"))
End Sub

Sub TestDictionary_Empty
  Dim dict
  Set dict = Dictionary(Array())
  AssertEqual 0, dict.Count
End Sub

Sub TestDictionary_Empty_shortcut
  Dim dict
  Set dict = D(Array())
  AssertEqual 0, dict.Count
End Sub

Sub TestDictionaryMerge
  Dim dict
  Set dict = DictionaryMerge(D(Array("foo", "Apple", "bar", "Banana")), D(Array("baz", "Orange")))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionaryMerge_shortcut
  Dim dict
  Set dict = DMerge(D(Array("foo", "Apple", "bar", "Banana")), D(Array("baz", "Orange")))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionaryMerge_OverwriteEntry
  Dim dict
  Set dict = DictionaryMerge(D(Array("foo", "Apple", "bar", "Banana")), D(Array("bar", "Pineapple", "baz", "Orange")))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Pineapple", dict("bar")
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionaryMerge_OverwriteEntry_shortcut
  Dim dict
  Set dict = DMerge(D(Array("foo", "Apple", "bar", "Banana")), D(Array("bar", "Pineapple", "baz", "Orange")))
  AssertEqual 3, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Pineapple", dict("bar")
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionaryMerge_EmptyLeft
  Dim dict
  Set dict = DictionaryMerge(D(Array("foo", "Apple", "bar", "Banana")), D(Array()))
  AssertEqual 2, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
End Sub

Sub TestDictionaryMerge_EmptyLeft_shortcut
  Dim dict
  Set dict = DMerge(D(Array("foo", "Apple", "bar", "Banana")), D(Array()))
  AssertEqual 2, dict.Count
  AssertEqual "Apple", dict("foo")
  AssertEqual "Banana", dict("bar")
End Sub

Sub TestDictionaryMerge_EmptyRight
  Dim dict
  Set dict = DictionaryMerge(D(Array()), D(Array("baz", "Orange")))
  AssertEqual 1, dict.Count
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionaryMerge_EmptyRight_shortcut
  Dim dict
  Set dict = DMerge(D(Array()), D(Array("baz", "Orange")))
  AssertEqual 1, dict.Count
  AssertEqual "Orange", dict("baz")
End Sub

Sub TestDictionaryMerge_EmptyBoth
  Dim dict
  Set dict = DictionaryMerge(D(Array()), D(Array()))
  AssertEqual 0, dict.Count
End Sub

Sub TestDictionaryMerge_EmptyBoth_shortcut
  Dim dict
  Set dict = DMerge(D(Array()), D(Array()))
  AssertEqual 0, dict.Count
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
