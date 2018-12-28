Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class MainFrm
    Public Const MAX_SEGMENT As Long = 1000
    Public Const MAX_SENTENCE As Long = 5000

    Dim g_objWord As Word.Application
    Dim g_objDoc As Word.Document

    Dim g_FontMain As New Font("맑은 고딕", 12, FontStyle.Regular)
    Dim g_FontSentence As New Font("맑은 고딕", 11, FontStyle.Regular)

    Dim g_NumSTPhrases As Long = 0
    Dim g_s_STPhrases(MAX_SENTENCE) As String
    Dim g_s_Segments(MAX_SEGMENT) As String

    Dim g_sFN_HTML As String = Nothing
    Dim g_sDIR_HTML As String = Nothing
    Dim g_sFN_DOCX_T As String = Nothing

    Public Function DeleteSegment(ByVal Content As String, ByVal sSegName As String) As String
        DeleteSegment = Nothing

        For i = 1 To Content.Length
            If Mid(Content, i, sSegName.Length + 1) = "<" & sSegName Then
                For j = i + 1 To Content.Length
                    If Mid(Content, j, 1) = ">" Then
                        DeleteSegment = Mid(Content, j + 1, Content.Length)
                        Exit For
                    End If
                Next

                Exit For
            End If
        Next

    End Function

    Public Sub ParseHTMLIntoPhrases()

        If FileIO.FileSystem.FileExists(g_sFN_HTML) = False Then Exit Sub

        Dim byString As String = FileIO.FileSystem.ReadAllText(g_sFN_HTML, System.Text.Encoding.Default)

        For i = 1 To byString.Length
            If Mid(byString, i, 5) = "<body" Then
                byString = Mid(byString, i, byString.Length)
                Exit For
            End If
        Next

        For i = 1 To byString.Length
            If Mid(byString, i, 6) = "</body" Then
                byString = Mid(byString, 1, i - 1)
                Exit For
            End If
        Next

        byString = DeleteSegment(byString, "body")
        byString = DeleteSegment(byString, "div")

        byString = byString.Replace("</div>", "")

        '### 이제부터가 파싱 시작! ###
        Dim nNumSegments As Long = 0
        Dim sSegments(MAX_SENTENCE) As String

        For i = 1 To byString.Length
            If Mid(byString, i, 2) = "<p" Then
                nNumSegments = nNumSegments + 1
                For j = i + 1 To byString.Length
                    If Mid(byString, j, 4) = "</p>" Then
                        sSegments(nNumSegments - 1) = Mid(byString, i, j - i + 4)
                        g_s_Segments(nNumSegments - 1) = sSegments(nNumSegments - 1) '### 전역 변수에 백업!!
                        Exit For
                    End If
                Next
            End If
        Next

        If nNumSegments > 0 Then
            g_NumSTPhrases = 0

            For i = 0 To nNumSegments - 1
                For j = 1 To sSegments(i).Length
                    If Mid(sSegments(i), j, 1) = "<" Then
                        For k = j + 1 To sSegments(i).Length
                            If Mid(sSegments(i), k, 1) = ">" Then
                                sSegments(i) = Mid(sSegments(i), k + 1, sSegments(i).Length)
                                j = 0
                                Exit For
                            End If
                        Next
                    Else
                        If sSegments(i).Length > 0 Then
                            g_NumSTPhrases = g_NumSTPhrases + 1

                            For k = j + 1 To sSegments(i).Length
                                If Mid(sSegments(i), k, 1) = "<" Then
                                    g_s_STPhrases(g_NumSTPhrases - 1) = Mid(sSegments(i), j, k - j)

                                    If g_s_STPhrases(g_NumSTPhrases - 1) = "&nbsp;" Then
                                        g_s_STPhrases(g_NumSTPhrases - 1) = Nothing
                                        g_NumSTPhrases = g_NumSTPhrases - 1
                                    End If

                                    j = k - 1
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Next
            Next
        End If


        LVST.Items.Clear()
        If g_NumSTPhrases > 0 Then
            For i = 0 To g_NumSTPhrases - 1
                g_s_STPhrases(i) = Net.WebUtility.HtmlDecode(g_s_STPhrases(i))
                LVST.Items.Add(New ListViewItem({"", i + 1, g_s_STPhrases(i), ""}))
            Next
        End If

    End Sub

    Public Sub DeleteFileAndDir()

        If g_sFN_HTML <> "" Then
            If My.Computer.FileSystem.FileExists(g_sFN_HTML) = True Then
                My.Computer.FileSystem.DeleteFile(g_sFN_HTML)
            End If
        End If

        If g_sDIR_HTML <> "" Then
            If My.Computer.FileSystem.DirectoryExists(g_sDIR_HTML) = True Then
                My.Computer.FileSystem.DeleteDirectory(g_sDIR_HTML, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
        End If

    End Sub

    Private Sub MainFrm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        g_objDoc = Nothing
        g_objWord = Nothing

        DeleteFileAndDir()

        End

    End Sub

    Private Sub InitiateVariables()
        '### 변수 초기화★ ###

        g_NumSTPhrases = 0

        For i = 0 To MAX_SENTENCE
            g_s_STPhrases(i) = Nothing
        Next

        For i = 0 To MAX_SEGMENT
            g_s_Segments(i) = Nothing
        Next

    End Sub

    Private Sub MainFrm_Load(sender As Object, e As EventArgs) Handles Me.Load

        TBSTSentence.Font = g_FontSentence
        TBTTSentence.Font = g_FontSentence

        LVST.View = View.Details
        LVST.FullRowSelect = True
        LVST.MultiSelect = False
        LVST.Columns.Add("")
        LVST.Columns.Add("번호")
        LVST.Columns.Add("원어")
        LVST.Columns.Add("번역")

        LVST.Columns(0).Width = 0
        LVST.Columns(1).Width = 40
        LVST.Columns(1).TextAlign = HorizontalAlignment.Center
        LVST.Columns(2).Width = 160
        LVST.Columns(3).Width = 160

        InitiateVariables()
    End Sub

    Private Sub 열기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 열기ToolStripMenuItem.Click
        OFD.FileName = Nothing
        OFD.Title = "파일 열기"
        OFD.Filter = "워드 파일(*.docx)|*.docx|텍스트 파일(*.txt)|*.txt"
        OFD.ShowDialog()

        If OFD.FileName IsNot "" Then

            InitiateVariables()

            '### 기존에 열었던 파일의 보조 파일은 지우자 ###
            DeleteFileAndDir()

            If Strings.Right(OFD.FileName, 3) = "doc" Or Strings.Right(OFD.FileName, 4) = "docx" Then

                g_sFN_HTML = OFD.FileName
                For i = 1 To g_sFN_HTML.Length
                    If Mid(g_sFN_HTML, i, 1) = "." Then
                        g_sFN_HTML = Strings.Left(g_sFN_HTML, i - 1)
                        g_sFN_DOCX_T = g_sFN_HTML & "_T.docx"
                        g_sDIR_HTML = g_sFN_HTML & ".files"
                        g_sFN_HTML = g_sFN_HTML & ".html"
                        Exit For
                    End If
                Next

                '### 똑같은 이름으로 기존에 있는 파일은 지우자 ###
                DeleteFileAndDir()

                Try
                    g_objWord = New Word.Application
                    g_objDoc = New Word.Document

                    g_objDoc = g_objWord.Documents.Add(OFD.FileName)
                    g_objDoc.SaveAs2(g_sFN_HTML, FileFormat:=10) '### 10 = Filtered HTML format ###
                    g_objDoc.Close()
                    g_objWord.Quit()

                    g_objDoc = Nothing
                    g_objWord = Nothing

                    WBTT.Navigate(g_sFN_HTML)

                Catch ex As Exception
                    MsgBox("오류")
                End Try

                ParseHTMLIntoPhrases()

            ElseIf Strings.Right(OFD.FileName, 3) = "txt" Then
                Dim byBuffer() As Byte = FileIO.FileSystem.ReadAllBytes(OFD.FileName)

                If byBuffer(0) = &HEF And byBuffer(1) = &HBB And byBuffer(2) = &HBF Then
                    'UTF8 파일이다
                    WBTT.Navigate(OFD.FileName)
                Else
                    '일반 텍스트 파일
                    WBTT.Navigate(OFD.FileName)
                End If

            End If


        End If

    End Sub

    Private Sub TSBFont_Click(sender As Object, e As EventArgs)

        FD.ShowDialog()

    End Sub

    Private Sub 저장ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 저장ToolStripMenuItem.Click
        '### 수정 후 HTML 파일을 열어서 DOCX로 변환!★ ###
        g_objDoc = g_objWord.Documents.Add(g_sFN_HTML)
        g_objDoc.SaveAs2(g_sFN_DOCX_T, FileFormat:=16) '### 16 = DOCX format ###
    End Sub

End Class

