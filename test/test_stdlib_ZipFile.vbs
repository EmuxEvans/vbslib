' stdlib.vbs: ZipFile class test.
' @import ../lib/stdlib.vbs

Option Explicit

Dim fso
Dim zip
Dim tempFolder

Sub SetUp
  Set fso = CreateObject("Scripting.FileSystemObject")
  tempFolder = fso.GetAbsolutePathName(".testZipFile")
  If fso.FolderExists(tempFolder) Then ' for debug
    fso.DeleteFolder(tempFolder)
  End If
  fso.CreateFolder(tempFolder)
  Set zip = ZipFile_Open(fso.BuildPath(tempFolder, "foo.zip"))
  zip.Timeout = 10
End Sub

Sub TearDown
  Set zip = Nothing
  fso.DeleteFolder(tempFolder)
  tempFolder = Empty
  Set fso = Nothing
End Sub

Sub TestZipFile_Empty
  Dim zipPath
  zipPath = fso.BuildPath(tempFolder, "foo.zip")
  Assert fso.FileExists(zipPath)
  AssertEqual Len(ZipFile_EmptyData), fso.GetFile(zipPath).Size
End Sub

Sub TestCopyHereAndItem_File
  With fso.OpenTextFile(fso.BuildPath(tempFolder, "bar.txt"), 2, True)
    .Write "Hello world."
    .Close
  End With
  zip.CopyHere fso.BuildPath(tempFolder, "bar.txt")

  Dim item
  Set item = zip.Item("bar.txt")
  Assert Not item.IsFolder
  AssertMatch "bar\.txt", item.Path
  AssertEqual Len("Hello world."), item.Size
  AssertMatch "foo\.zip", item.Parent.Title
End Sub

Sub TestCopyHereAndItem_Folder
  fso.CreateFolder(fso.BuildPath(tempFolder, "bar"))
  ' need for folder contents
  With fso.OpenTextFile(fso.BuildPath(tempFolder, "bar\baz.txt"), 2, True)
    .Write "Hello world."
    .Close
  End With
  zip.CopyHere fso.BuildPath(tempFolder, "bar")

  Dim item
  Set item = zip.Item("bar")
  Assert item.IsFolder
  AssertMatch "bar", item.Path
  AssertMatch "foo\.zip", item.Parent.Title
End Sub

Sub TestItem_NotFound
  Dim errNum, errSrc
  On Error Resume Next
  zip.Item("bar")
  errNum = Err.Number
  errSrc = Err.Source
  Err.Clear
  On Error GoTo 0

  AssertEqual 51, errNum
  AssertEqual "stdlib.vbs:ZipFile.Item(Get)", errSrc
End Sub

Sub TestSubFolder
  fso.CreateFolder(fso.BuildPath(tempFolder, "bar"))
  ' need for folder contents
  With fso.OpenTextFile(fso.BuildPath(tempFolder, "bar\baz.txt"), 2, True)
    .Write "Hello world."
    .Close
  End With
  zip.CopyHere fso.BuildPath(tempFolder, "bar")

  Dim item
  Set item = zip.SubFolder("bar").Item("baz.txt")
  Assert Not item.IsFolder
  AssertMatch "baz\.txt", item.Path
  AssertEqual Len("Hello world."), item.Size
  AssertMatch "bar", item.Parent
  AssertMatch "foo\.zip", item.Parent.ParentFolder.Title
End Sub

Sub TestSubFolder_NotFound
  Dim errNum, errSrc, errDsc
  On Error Resume Next
  zip.SubFolder("bar")
  errNum = Err.Number
  errSrc = Err.Source
  errDsc = Err.Description
  Err.Clear
  On Error GoTo 0

  AssertEqual 51, errNum
  AssertEqual "stdlib.vbs:ZipFile.SubFolder(Get)", errSrc
  AssertMatch re("not found", "i"), errDsc
End Sub

Sub TestSubFolder_NotFolder
  With fso.OpenTextFile(fso.BuildPath(tempFolder, "bar.txt"), 2, True)
    .Write "Hello world."
    .Close
  End With
  zip.CopyHere fso.BuildPath(tempFolder, "bar.txt")

  Dim errNum, errSrc, errDsc
  On Error Resume Next
  zip.SubFolder("bar.txt")
  errNum = Err.Number
  errSrc = Err.Source
  errDsc = Err.Description
  Err.Clear
  On Error GoTo 0

  AssertEqual 51, errNum
  AssertEqual "stdlib.vbs:ZipFile.SubFolder(Get)", errSrc
  AssertMatch re("not.*folder", "i"), errDsc
End Sub

Sub TestItems
  fso.CreateFolder(fso.BuildPath(tempFolder, "bar"))
  ' need for folder contents
  With fso.OpenTextFile(fso.BuildPath(tempFolder, "bar\baz.txt"), 2, True)
    .Write "Hello world."
    .Close
  End With
  With fso.OpenTextFile(fso.BuildPath(tempFolder, "quux.txt"), 2, True)
    .Write "Hello world."
    .Close
  End With
  zip.CopyHere fso.BuildPath(tempFolder, "bar")
  zip.CopyHere fso.BuildPath(tempFolder, "quux.txt")

  Dim count, item
  count = 0
  For Each item In zip.Items
    Select Case Count
      Case 0:
        Assert item.IsFolder
        AssertMatch "bar", item.Path
        AssertMatch "foo\.zip", item.Parent.Title
      Case 1:
        Assert Not item.IsFolder
        AssertMatch "quux\.txt", item.Path
        AssertEqual Len("Hello world."), item.Size
        AssertMatch "foo\.zip", item.Parent.Title
      Case Else:
    End Select
    count = count + 1
  Next
  AssertEqual 2, count
End Sub

Sub TestItems_Emtpy
  AssertEqual 0, CountItem(zip.Items)
End Sub

Sub TestSubFolders_Empty
  AssertEqual 0, CountItem(zip.SubFolders)
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
