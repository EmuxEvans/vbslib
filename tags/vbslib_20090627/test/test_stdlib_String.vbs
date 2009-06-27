' stdlib.vbs: string formatting test.
' @import ../lib/stdlib.vbs

Option Explicit

Sub TestLPad_Padding
  AssertEqual "  foo", LPad("foo", 5, " ")
End Sub

Sub TestLPad_NoPadding
  AssertEqual "foo", LPad("foo", 3, " ")
End Sub

Sub TestLPad_NoPadding2
  AssertEqual "foo", LPad("foo", 1, " ")
End Sub

Sub TestRPad_Padding
  AssertEqual "foo  ", RPad("foo", 5, " ")
End Sub

Sub TestRPad_NoPadding
  AssertEqual "foo", RPad("foo", 3, " ")
End Sub

Sub TestRPad_NoPadding2
  AssertEqual "foo", RPad("foo", 1, " ")
End Sub

Sub TestStrftime_EscapeSpecialCharacter
  AssertEqual "%", strftime("%%", Now)
End Sub

Sub TestStrftime_EscapeSpecialCharacterWithPrefix
  AssertEqual "foo %", strftime("foo %%", Now)
End Sub

Sub TestStrftime_EscapeSpecialCharacterWithSuffix
  AssertEqual "% foo", strftime("%% foo", Now)
End Sub

Sub TestStrftime_YyyyMmDdHhMmSs
  AssertEqual "2009-01-02 03:04:05", strftime("%Y-%m-%d %H:%M:%S", #2009-01-02 03:04:05#)
End Sub

Sub TestStrftime_YyyyMmDdHhMmSsWithPrefixSuffix
  AssertEqual "[2009-01-02 03:04:05]", strftime("[%Y-%m-%d %H:%M:%S]", #2009-01-02 03:04:05#)
End Sub

Sub TestStrftime_YyyyMmDdHhMmSsNoSeparator
  AssertEqual "20090102030405", strftime("%Y%m%d%H%M%S", #2009-01-02 03:04:05#)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
