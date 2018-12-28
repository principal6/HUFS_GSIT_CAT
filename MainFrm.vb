Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class MainFrm
    Public Const MAX_HTMLSEGMENT As Long = 1000
    Public Const MAX_HTMLSUBSEGMENT As Long = 3000
    Public Const MAX_SENTENCE As Long = 5000

    Dim g_objWord As Word.Application
    Dim g_objDoc As Word.Document

    Dim g_FontSentence As New Font("맑은 고딕", 10, FontStyle.Regular)

    Dim g_sHTMLSegs(MAX_HTMLSEGMENT) As String
    Dim g_nHTMLSegsCount As Long = 0

    Public Structure SUBSEG
        Dim Content As String
        Dim ParentSegID As Long
    End Structure

    Public Structure PHRASES
        Dim Content As String
        Dim HTMLSubSegID As Long
        Dim HTMLSegID As Long
    End Structure

    Dim g_sHTMLSubSegs(MAX_HTMLSUBSEGMENT) As SUBSEG
    Dim g_nHTMLSubSegsCount As Long = 0
    Dim g_sSTPhrases(MAX_SENTENCE) As PHRASES
    Dim g_nSTPhrasesCount As Long = 0

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
        g_nHTMLSegsCount = 0
        For i = 1 To byString.Length
            If Mid(byString, i, 2) = "<p" Then
                g_nHTMLSegsCount = g_nHTMLSegsCount + 1
                For j = i + 1 To byString.Length
                    If Mid(byString, j, 4) = "</p>" Then
                        g_sHTMLSegs(g_nHTMLSegsCount - 1) = Mid(byString, i, j - i + 4)
                        g_sHTMLSegs(g_nHTMLSegsCount - 1) = g_sHTMLSegs(g_nHTMLSegsCount - 1).Replace(vbCrLf, " ")
                        Exit For
                    End If
                Next
            End If
        Next

        If g_nHTMLSegsCount > 0 Then

            For i = 0 To g_nHTMLSegsCount - 1
                For j = 1 To g_sHTMLSegs(i).Length
                    If Mid(g_sHTMLSegs(i), j, 1) = "<" Then
                        For k = j + 1 To g_sHTMLSegs(i).Length
                            If Mid(g_sHTMLSegs(i), k, 1) = ">" Then
                                j = k
                                Exit For
                            End If
                        Next
                    Else
                        For k = j + 1 To g_sHTMLSegs(i).Length
                            If Mid(g_sHTMLSegs(i), k, 1) = "<" Then
                                g_nHTMLSubSegsCount = g_nHTMLSubSegsCount + 1
                                g_sHTMLSubSegs(g_nHTMLSubSegsCount - 1).Content = Mid(g_sHTMLSegs(i), j, k - j)
                                g_sHTMLSubSegs(g_nHTMLSubSegsCount - 1).ParentSegID = i
                                j = k - 1
                                Exit For
                            End If
                        Next
                    End If
                Next
            Next


            '### .을 기준으로 문장 구분하자 ###
            g_nSTPhrasesCount = 0

            For i = 0 To g_nHTMLSubSegsCount - 1
                For j = 1 To g_sHTMLSubSegs(i).Content.Length
                    If Mid(g_sHTMLSubSegs(i).Content, j, 1) = "." Then

                        If j = 1 Then
                            '### 맨 처음인데 .이 찍혀있으면 나누지 않음
                        ElseIf Mid(g_sHTMLSubSegs(i).Content, j - 1, 1) = ">" Or Mid(g_sHTMLSubSegs(i).Content, j - 1, 1) = "." Then
                            '### 맨 처음 글자가 .이거나 바로 앞 글자도 .이었을 경우 나누지 않음
                        Else
                            g_nSTPhrasesCount = g_nSTPhrasesCount + 1
                            g_sSTPhrases(g_nSTPhrasesCount - 1).Content = Mid(g_sHTMLSubSegs(i).Content, 1, j)
                            g_sSTPhrases(g_nSTPhrasesCount - 1).HTMLSubSegID = i
                            g_sSTPhrases(g_nSTPhrasesCount - 1).HTMLSegID = g_sHTMLSubSegs(g_sSTPhrases(g_nSTPhrasesCount - 1).HTMLSubSegID).ParentSegID
                            g_sHTMLSubSegs(i).Content = Trim(Mid(g_sHTMLSubSegs(i).Content, j + 1, g_sHTMLSubSegs(i).Content.Length))
                            j = 0
                        End If
                    End If
                Next

                g_sHTMLSubSegs(i).Content = g_sHTMLSubSegs(i).Content.Replace("&nbsp;", "") '### &nbsp; 는 빈 내용이므로 번역이 필요 없으니까!!

                If g_sHTMLSubSegs(i).Content <> "" Then '### 아직 내용이 남아 있으면 새로 추가!!
                    g_nSTPhrasesCount = g_nSTPhrasesCount + 1
                    g_sSTPhrases(g_nSTPhrasesCount - 1).Content = Trim(g_sHTMLSubSegs(i).Content)
                    g_sSTPhrases(g_nSTPhrasesCount - 1).HTMLSubSegID = i
                    g_sSTPhrases(g_nSTPhrasesCount - 1).HTMLSegID = g_sHTMLSubSegs(g_sSTPhrases(g_nSTPhrasesCount - 1).HTMLSubSegID).ParentSegID
                End If
            Next

        End If

        LVST.Items.Clear()
        If g_nSTPhrasesCount > 0 Then
            For i = 0 To g_nSTPhrasesCount - 1
                g_sSTPhrases(i).Content = Net.WebUtility.HtmlDecode(g_sSTPhrases(i).Content)
                'LVST.Items.Add(New ListViewItem({"", i + 1, g_sSTPhrases(i).Content, g_sSTPhrases(i).HTMLSegID})) '### 디버그용
                LVST.Items.Add(New ListViewItem({"", i + 1, g_sSTPhrases(i).Content, ""}))
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

        g_nHTMLSegsCount = 0
        For i = 0 To MAX_HTMLSEGMENT
            g_sHTMLSegs(i) = Nothing
        Next

        g_nHTMLSubSegsCount = 0
        For i = 0 To MAX_HTMLSUBSEGMENT
            g_sHTMLSubSegs(i).Content = Nothing
            g_sHTMLSubSegs(i).ParentSegID = 0
        Next

        g_nSTPhrasesCount = 0
        For i = 0 To MAX_SENTENCE
            g_sSTPhrases(i).Content = Nothing
            g_sSTPhrases(i).HTMLSegID = 0
            g_sSTPhrases(i).HTMLSubSegID = 0
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
        LVST.Columns(2).Width = 300
        LVST.Columns(3).Width = 280

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
        If g_sFN_HTML = Nothing Then Exit Sub

        g_objDoc = g_objWord.Documents.Add(g_sFN_HTML)
        g_objDoc.SaveAs2(g_sFN_DOCX_T, FileFormat:=16) '### 16 = DOCX format ###
    End Sub

    Private Sub LVST_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LVST.SelectedIndexChanged
        If LVST.FocusedItem IsNot Nothing Then
            TBSTSentence.Text = g_sSTPhrases(LVST.FocusedItem.Index).Content
        End If
    End Sub
End Class

