Module ModFile
    Public Function SAFE_DELETE_FILE(ByVal FileName As String) As Boolean
        SAFE_DELETE_FILE = False

        If FileName = Nothing Then Exit Function

        If My.Computer.FileSystem.FileExists(FileName) = True Then
            My.Computer.FileSystem.DeleteFile(FileName)
        End If

        SAFE_DELETE_FILE = True
    End Function

    Public Function SAFE_DELETE_DIR(ByVal Dir As String) As Boolean
        SAFE_DELETE_DIR = False

        If Dir = Nothing Then Exit Function

        If My.Computer.FileSystem.DirectoryExists(Dir) = True Then
            My.Computer.FileSystem.DeleteDirectory(Dir, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If

        SAFE_DELETE_DIR = True
    End Function
End Module
