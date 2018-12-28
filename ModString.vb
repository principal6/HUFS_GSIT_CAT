Module ModString

    Public Function String_ExistSpace(ByVal Text As String) As Boolean
        String_ExistSpace = False

        If Text = Nothing Then Exit Function

        For i = 1 To Text.Length
            If Mid(Text, i, 1) = " " Then
                String_ExistSpace = True
                Exit Function
            End If
        Next
    End Function

    Public Function String_ReplaceByPos(ByVal Text As String, ByVal StartPos As Integer, ByVal EndPos As Integer,
        ByVal ReplaceString As String) As String

        Dim sOutput As String = Nothing

        Dim T_sLeft As String = Nothing
        Dim T_sRight As String = Nothing

        T_sLeft = Strings.Left(Text, StartPos - 1)
        T_sRight = Strings.Mid(Text, EndPos + 1, Text.Length)
        sOutput = T_sLeft & ReplaceString & T_sRight

        String_ReplaceByPos = sOutput
    End Function

    Public Function String_Cut_Ext(ByVal FileName As String) As String
        String_Cut_Ext = Nothing

        If FileName = Nothing Then Exit Function

        Dim Output As String = Nothing
        For i = FileName.Length To 1 Step -1
            If Mid(FileName, i, 1) = "." Then
                Output = Strings.Left(FileName, i - 1)
                Exit For
            End If
        Next

        String_Cut_Ext = Output
    End Function
End Module
