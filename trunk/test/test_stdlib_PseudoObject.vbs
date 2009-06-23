' stdlib.vbs: Pseudo object test.
' @import ../lib/stdlib.vbs

Option Explicit

Function Counter_Create
  Dim c
  Set c = D(Array("count", 0))
  PseudoObject_AttachMethodSubProc c, "CountUp", GetRef("Counter_CountUp"), 1
  PseudoObject_AttachMethodFuncProc c, "GetValue", GetRef("Counter_GetValue"), 1
  PseudoObject_AttachMethodSubProc c, "SetValue", GetRef("Counter_SetValue"), 2
  Set Counter_Create = c
End Function

Sub Counter_CountUp(self)
  self("count") = self("count") + 1
End Sub

Function Counter_GetValue(self)
  Counter_GetValue = self("count")
End Function

Sub Counter_SetValue(self, newValue)
  self("count") = newValue
End Sub

Dim counter

Sub SetUp
  Set counter = Counter_Create
End Sub

Sub TearDown
  Set counter = Nothing
End Sub

Sub TestPseudoObjectMethodCall
  AssertEqual 0, counter("GetValue")()

  counter("CountUp")()
  AssertEqual 1, counter("GetValue")()

  counter("CountUp")()
  AssertEqual 2, counter("GetValue")()

  counter("CountUp")()
  AssertEqual 3, counter("GetValue")()

  counter("SetValue")(100)
  AssertEqual 100, counter("GetValue")()

  counter("CountUp")()
  AssertEqual 101, counter("GetValue")()
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
