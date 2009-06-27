' file open error test
' @import ../lib/stdlib.vbs

Dim fso
Dim tempFolder

' for fso.OpenTextFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Sub SetUp
  Set fso = CreateObject("Scripting.FileSystemObject")
  tempFolder = "temp_FileOpenError"
  fso.CreateFolder tempFolder
End Sub

Sub TearDown
  fso.DeleteFolder tempFolder
  Set fso = Nothing
End Sub

Sub TestFileSystemObject_CreateTextFile_OverwriteFailed
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))
    .Close
  End With

  Dim errNum
  On Error Resume Next
  fso.CreateTextFile fso.BuildPath(tempFolder, "foo.txt"), False
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  AssertEqual 58, errNum
End Sub

Sub TestFileSystemObject_CreateTextFile_WriteOpenFailed
  Dim f
  Set f = fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))

  Dim errNum
  On Error Resume Next
  fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  f.Close

  AssertEqual 70, errNum
End Sub

Sub TestFileSystemObject_OpenTextFile_NotFileExists
  Dim errNum
  On Error Resume Next
  fso.OpenTextFile fso.BuildPath(tempFolder, "foo.txt")
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  AssertEqual 53, errNum
End Sub

Sub TestFileSystemObject_OpenTextFile_WriteOpenFailed
  Dim f
  Set f = fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))

  Dim errNum
  On Error Resume Next
  fso.OpenTextFile fso.BuildPath(tempFolder, "foo.txt"), ForWriting
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  f.Close

  AssertEqual 70, errNum
End Sub

Sub TestFileSystemObject_OpenTextFile_AppendOpenFailed
  Dim f
  Set f = fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))

  Dim errNum
  On Error Resume Next
  fso.OpenTextFile fso.BuildPath(tempFolder, "foo.txt"), ForAppending
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  f.Close

  AssertEqual 70, errNum
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
