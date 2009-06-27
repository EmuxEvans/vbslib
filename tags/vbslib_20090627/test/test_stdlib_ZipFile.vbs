' stdlib.vbs: ZipFile class test.
' @import ../lib/stdlib.vbs

Option Explicit

Dim shell
Dim fso
Dim zfo
Dim tempFolder

' for fso.OpenTextFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Function GetWindowsName
  GetWindowsName = FirstItem(GetObject("winmgmts:root\cimv2").InstancesOf("Win32_OperatingSystem")).Caption
End Function

Sub SetUp
  Set shell = CreateObject("WScript.Shell")
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set zfo = New ZipFileObject
  zfo.TimeoutSeconds = 1
  tempFolder = "temp_ZipFile"
  fso.CreateFolder tempFolder
End Sub

Sub TearDown
  fso.DeleteFolder tempFolder
  Set shell = Nothing
  Set fso = Nothing
  Set zfo = Nothing
End Sub

Sub TestTimeoutSeconds_Default
  Set zfo = New ZipFileObject
  AssertEqual 60, zfo.TimeoutSeconds
End Sub

Sub TestTimeoutSeoncs_LetValue
  zfo.TimeoutSeconds = 100
  AssertEqual 100, zfo.TimeoutSeconds
End Sub

Sub TestPollingMillisecs_Default
  Set zfo = New ZipFileObject
  AssertEqual 100, zfo.PollingIntervalMillisecs
End Sub

Sub TestPollingMillisecs_LetValue
  zfo.PollingIntervalMillisecs = 123
  AssertEqual 123, zfo.PollingIntervalMillisecs
End Sub

Sub TestIsOpened_Opened
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))
    Assert zfo.IsOpened(fso.BuildPath(tempFolder, "foo.txt"))
    .Close
  End With
End Sub

Sub TestIsOpened_NotOpened
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))
    .Close
  End With

  Assert Not zfo.IsOpened(fso.BuildPath(tempFolder, "foo.txt"))
End Sub

Sub TestIsOpened_NotFileExists
  Dim errNum

  On Error Resume Next
  zfo.IsOpened(fso.BuildPath(tempFolder, "foo.txt"))
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  AssertEqual 53, errNum
End Sub

Sub TestCreateEmptyZipFile
  zfo.CreateEmptyZipFile fso.BuildPath(tempFolder, "foo.zip"), False

  Assert fso.FileExists(fso.BuildPath(tempFolder, "foo.zip"))
  With fso.GetFile(fso.BuildPath(tempFolder, "foo.zip"))
    AssertEqual Len(ZipFile_EmptyData), .Size
    With .OpenAsTextStream(ForReading)
      AssertEqual ZipFile_EmptyData, .ReadAll
      .Close
    End With
  End With
End Sub

Sub TestCreateEmptyZipFile_Overwrite
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.zip"))
    .Write "foo"
    .Close
  End With

  zfo.CreateEmptyZipFile fso.BuildPath(tempFolder, "foo.zip"), True

  Assert fso.FileExists(fso.BuildPath(tempFolder, "foo.zip"))
  With fso.GetFile(fso.BuildPath(tempFolder, "foo.zip"))
    AssertEqual Len(ZipFile_EmptyData), .Size
    With .OpenAsTextStream(ForReading)
      AssertEqual ZipFile_EmptyData, .ReadAll
      .Close
    End With
  End With
End Sub

Sub TestCreateEmptyZipFile_NotOverwrite
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.zip"))
    .Write "foo"
    .Close
  End With

  Dim errNum

  On Error Resume Next
  zfo.CreateEmptyZipFile fso.BuildPath(tempFolder, "foo.zip"), False
  errNum = Err.Number
  Err.Clear
  On Error GoTo 0

  AssertEqual 58, errNum
End Sub

Sub TestZip
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))
    .Write "Hello world."
    .Close
  End With

  zfo.Zip fso.BuildPath(tempFolder, "foo.zip"), _
     Array(fso.BuildPath(tempFolder, "foo.txt"))

  Assert fso.FileExists(fso.BuildPath(tempFolder, "foo.zip"))
  With fso.GetFile(fso.BuildPath(tempFolder, "foo.zip"))
    Assert .Size > Len(ZipFile_EmptyData)
  End With
End Sub

Sub TestZipAndUnzip
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))
    .Write "Hello world."
    .Close
  End With

  Dim zipPath
  zipPath = fso.BuildPath(tempFolder, "foo.zip")
  zfo.Zip zipPath, _
     Array(fso.BuildPath(tempFolder, "foo.txt"))

  Dim extractPath
  extractPath = fso.BuildPath(tempFolder, "extract")
  fso.CreateFolder extractPath
  zfo.Unzip zipPath, extractPath

  Assert fso.FileExists(fso.BuildPath(extractPath, "foo.txt"))
  With fso.GetFile(fso.BuildPath(extractPath, "foo.txt"))
    AssertEqual Len("Hello world."), .Size
    With .OpenAsTextStream(ForReading)
      AssertEqual "Hello world.", .ReadAll
      .Close
    End With
  End With

  ' Why does unzip temporary file remain in temporary directory at Windows XP?
  If re("Windows XP", "i").Test(GetWindowsName) Then
    Dim systemTempFolder
    Set systemTempFolder = fso.GetFolder(shell.ExpandEnvironmentStrings("%TEMP%"))

    Dim delItems
    delItems = FindAll(systemTempFolder.SubFolders, _
                       ValueFilter(ValueObjectProperty("Name"), _
                                   ValueMatch(re("foo\.zip", "i"))))

    Dim i
    For Each i In delItems
      i.Delete(True)
    Next
  End If
End Sub

Sub TestZipAndZipEntries
  With fso.CreateTextFile(fso.BuildPath(tempFolder, "foo.txt"))
    .Write "Hello world."
    .Close
  End With

  Dim zipPath
  zipPath = fso.BuildPath(tempFolder, "foo.zip")
  zfo.Zip zipPath, _
     Array(fso.BuildPath(tempFolder, "foo.txt"))

  AssertEqual ShowValue(Array(fso.BuildPath(zipPath, "foo.txt"))), _
              ShowValue(zfo.ZipEntries(zipPath))
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
