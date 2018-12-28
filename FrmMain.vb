Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class FrmMain
    Dim g_objWord As Word.Application
    Dim g_objDoc As Word.Document

    Dim g_FontSentence As New Font("맑은 고딕", 10, FontStyle.Regular)
    Dim g_nPrevLVSTIndex As Integer = 0

    Dim g_HTMLText As HTML_TEXT
    Dim g_HTMLText_T As HTML_TEXT

    Dim g_HTMLSegOrg(MAX_HTMLSEGMENT) As HTML_SEG
    Dim g_sHTMLSegTrs(MAX_HTMLSEGMENT) As String
    Dim g_nHTMLSegCount As Integer = 0
    Dim g_bHTMLSegTranslated(MAX_HTMLSEGMENT) As Boolean

    Dim g_HTMLSubSeg(MAX_HTMLSUBSEGMENT) As HTML_SUBSEG
    Dim g_nHTMLSubSegCount As Integer = 0

    Dim g_PhraseOrg(MAX_PHRASE) As HTML_PHRASE
    Dim g_nPhraseCount As Integer = 0
    Dim g_sPhraseTrs(MAX_PHRASE) As String

    '### HTML파일이 저장될 파일과 폴더
    Dim g_sDIR_HTML As String = Nothing
    Dim g_sFN_HTML As String = Nothing

    '### WebBrowser 스크롤바 위치 저장
    Dim g_WBScrollTop As Integer
    Dim g_WBScrollLeft As Integer

    Public Sub DeleteFileAndDir()
        SAFE_DELETE_FILE(g_sFN_HTML)
        SAFE_DELETE_DIR(g_sDIR_HTML)
    End Sub

    Public Sub HTML_ParseIntoPhrases()
        If FileIO.FileSystem.FileExists(g_sFN_HTML) = False Then Exit Sub

        Dim T_HTMLText As String = FileIO.FileSystem.ReadAllText(g_sFN_HTML, System.Text.Encoding.Default)
        g_HTMLText.ReadFromText(T_HTMLText)
        g_HTMLText_T = g_HTMLText

        '### HTML_PARSE_01: Seg 구하기 ###
        g_nHTMLSegCount = GetSegFromHTMLBody(g_HTMLText.Body, g_HTMLSegOrg)

        If g_nHTMLSegCount = 0 Then Exit Sub

        '### Seg에서 <span lang=~> 지우기: 각 HTML Segment 당 실행 (<p ~> ~ </p>)
        For i = 0 To g_nHTMLSegCount - 1
            HTML_DeleteSpanLang(g_HTMLSegOrg(i))
        Next

        '### HTML_PARSE_02: SubSeg 찾기: 각 HTML Segment 당 실행 (<p ~> ~ </p>)
        For i = 0 To g_nHTMLSegCount - 1
            GetSubSegFromSeg(i, g_HTMLSegOrg(i), g_HTMLSubSeg, g_nHTMLSubSegCount)
        Next

        '### HTML_PARSE_03: Phrase(문장) 나누기 ★ ###
        g_nPhraseCount = 0
        For i = 0 To g_nHTMLSubSegCount - 1
            GetPhraseFromSubSeg(i, g_HTMLSubSeg(i), g_PhraseOrg, g_nPhraseCount)
        Next

        '### ListView에 Phrase 채우기
        LVST.Items.Clear()
        If g_nPhraseCount > 0 Then
            For i = 0 To g_nPhraseCount - 1
                g_PhraseOrg(i).Content = Net.WebUtility.HtmlDecode(g_PhraseOrg(i).Content)
                LVST.Items.Add(New ListViewItem({"", i + 1, g_PhraseOrg(i).Content, ""}))
            Next
        End If

    End Sub


    Public Sub TranslatePhrase(ByVal CurItemID As Integer)
        Dim CurSegID As Integer = g_PhraseOrg(CurItemID).ParentSegID
        Dim CurSubSegRelID As Integer = g_HTMLSubSeg(g_PhraseOrg(CurItemID).ParentSubSegID).IndexInSeg
        Dim CurPhraseRelID As Integer = g_PhraseOrg(CurItemID).IndexInSubSeg

        g_sPhraseTrs(CurItemID) = TBTTSentence.Text

        Dim bPhraseDivided As Boolean = False

        If g_HTMLSubSeg(g_PhraseOrg(CurItemID).ParentSubSegID).PhraseCount > 1 Then
            '### 문장 구분이 있다!
            bPhraseDivided = True
        End If

        '### PhraseTrs에 아무 내용도 없으면? 번역하지 말기!
        If g_sPhraseTrs(CurItemID) = "" Then Exit Sub

        '### PhraseTrs에 내용이 있으니까 번역하기! ★
        Dim sTrgEncoded As String = HTML_Encode(g_sPhraseTrs(CurItemID))
        UpdateSegment(g_bHTMLSegTranslated(CurSegID), g_HTMLSegOrg(CurSegID), g_sHTMLSegTrs(CurSegID), CurSubSegRelID, CurPhraseRelID, bPhraseDivided, sTrgEncoded)

        '### ListView에 수정!
        LVST.Items(CurItemID).SubItems(3).Text = g_sPhraseTrs(CurItemID)

        '### 이제 HTML 수정!!
        UpdateHTMLText()
    End Sub

    Public Sub UpdateHTMLText()
        g_HTMLText_T.Body = Nothing

        For i = 0 To g_nHTMLSegCount - 1
            If g_bHTMLSegTranslated(i) = True Then
                g_HTMLText_T.Body = g_HTMLText_T.Body & g_sHTMLSegTrs(i)
            Else
                g_HTMLText_T.Body = g_HTMLText_T.Body & g_HTMLSegOrg(i).Content
            End If
        Next

        '### WebBrowser 스크롤 지정
        g_WBScrollTop = WBTT.Document.Body.ScrollTop
        g_WBScrollLeft = WBTT.Document.Body.ScrollLeft

        '### 바뀐 HTML을 저장한다
        FileIO.FileSystem.WriteAllText(g_sFN_HTML, g_HTMLText_T.GetText, False)
        WBTT.Navigate(g_sFN_HTML)
    End Sub

    Public Sub UntranslatePhrase(ByVal PhraseID As Integer)
        TBTTSentence.Text = g_PhraseOrg(PhraseID).Content
        TranslatePhrase(PhraseID)

        TBTTSentence.Text = Nothing
        g_sPhraseTrs(PhraseID) = Nothing
        LVST.Items(PhraseID).SubItems(3).Text = Nothing
    End Sub

    Private Sub InitializeVariables()
        '### 변수 초기화★ ###

        g_HTMLText.Clear()
        g_HTMLText_T = g_HTMLText

        g_nHTMLSegCount = 0
        For i = 0 To MAX_HTMLSEGMENT
            g_HTMLSegOrg(i).Clear()
        Next

        g_nHTMLSubSegCount = 0
        For i = 0 To MAX_HTMLSUBSEGMENT
            g_HTMLSubSeg(i).Clear()
        Next

        g_nPhraseCount = 0
        For i = 0 To MAX_PHRASE
            g_PhraseOrg(i).Clear()
            g_sPhraseTrs(i) = Nothing
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
        LVST.Columns(1).Width = 42
        LVST.Columns(1).TextAlign = HorizontalAlignment.Center
        LVST.Columns(2).Width = 300
        LVST.Columns(3).Width = 280

        LVST.ContextMenuStrip = CMS1
        InitializeVariables()
    End Sub

    Private Sub MainFrm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        g_objDoc = Nothing
        g_objWord = Nothing

        DeleteFileAndDir()

        End '### 프로그램 안전하게 끝내기★
    End Sub

    Private Sub 열기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 열기ToolStripMenuItem.Click
        OFD.FileName = Nothing
        OFD.Title = "파일 열기"
        OFD.Filter = "워드 파일(*.docx)|*.docx|텍스트 파일(*.txt)|*.txt"
        OFD.ShowDialog()

        If OFD.FileName IsNot "" Then

            InitializeVariables()

            '### 기존에 열었던 파일의 보조 파일은 지우자 ###
            DeleteFileAndDir()

            If Strings.Right(OFD.FileName, 3) = "doc" Or Strings.Right(OFD.FileName, 4) = "docx" Then

                g_sFN_HTML = OFD.FileName
                For i = 1 To g_sFN_HTML.Length
                    If Mid(g_sFN_HTML, i, 1) = "." Then
                        g_sFN_HTML = Strings.Left(g_sFN_HTML, i - 1)
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

                HTML_ParseIntoPhrases()

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

    Private Sub 저장ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 저장ToolStripMenuItem.Click
        '### 수정 후 HTML 파일을 열어서 DOCX로 변환!★ ###

        If g_sFN_HTML = Nothing Then Exit Sub

        SFD.FileName = Nothing
        SFD.Title = "파일 저장하기"
        SFD.Filter = "워드 파일(*.docx)|*.docx"
        SFD.ShowDialog()

        If SFD.FileName IsNot "" Then
            g_objWord = New Word.Application
            g_objDoc = New Word.Document

            g_objDoc = g_objWord.Documents.Add(g_sFN_HTML)

            Dim objInlineShape As Word.InlineShape

            For Each objInlineShape In g_objWord.ActiveDocument.InlineShapes
                '### 그림 파일 포함해서 저장!★★ ###
                objInlineShape.LinkFormat.SavePictureWithDocument = True
                objInlineShape.LinkFormat.BreakLink()
            Next

            g_objDoc.SaveAs2(SFD.FileName, FileFormat:=16) '### 16 = DOCX format ###

            g_objDoc.Close()
            g_objWord.Quit()

            g_objDoc = Nothing
            g_objWord = Nothing
        End If

    End Sub

    Private Sub LVST_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LVST.SelectedIndexChanged
        If LVST.FocusedItem IsNot Nothing Then
            TBSTSentence.Text = g_PhraseOrg(LVST.FocusedItem.Index).Content
            TBTTSentence.Text = g_sPhraseTrs(LVST.FocusedItem.Index)
        End If
    End Sub

    Private Sub LVST_KeyDown(sender As Object, e As KeyEventArgs) Handles LVST.KeyDown
        If e.KeyCode = Keys.Enter Then
            If LVST.Items.Count > 0 Then
                g_nPrevLVSTIndex = LVST.FocusedItem.Index
                TBTTSentence.Focus()
            End If
        End If
    End Sub

    Private Sub TBTTSentence_KeyDown(sender As Object, e As KeyEventArgs) Handles TBTTSentence.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True '엔터 키 눌림 방지

            If g_sPhraseTrs(g_nPrevLVSTIndex) <> TBTTSentence.Text Then
                '### 달라진 게 있으면 반영하자!

                If TBTTSentence.Text = "" Then
                    '### TBTTSentence.Text가 Null인 경우
                    TBTTSentence.Text = g_sPhraseTrs(g_nPrevLVSTIndex)
                Else
                    '### TBTTSentence.Text가 Null이 아닌 경우
                    TranslatePhrase(g_nPrevLVSTIndex)
                End If
            End If

            LVST.Items(g_nPrevLVSTIndex).Selected = True
            LVST.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            LVST.Items(g_nPrevLVSTIndex).Selected = True
            LVST.Focus()
        End If
    End Sub

    Private Sub TBTTSentence_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TBTTSentence.KeyPress
        If e.KeyChar = Convert.ToChar(1) Then
            TBTTSentence.SelectAll()
        End If
    End Sub

    Private Sub TBSTSentence_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TBSTSentence.KeyPress
        If e.KeyChar = Convert.ToChar(1) Then
            TBSTSentence.SelectAll()
        End If
    End Sub

    Private Sub 번역하기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 번역하기ToolStripMenuItem.Click
        If LVST.Items.Count > 0 Then
            g_nPrevLVSTIndex = LVST.FocusedItem.Index
            TBTTSentence.Focus()
        End If
    End Sub

    Private Sub 번역되돌리기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 번역되돌리기ToolStripMenuItem.Click
        If LVST.Items.Count > 0 Then
            If MsgBox("해당 번역을 지우고 원본으로 되돌리시겠습니까?", MsgBoxStyle.YesNo, "번역 되돌리기") = vbYes Then
                '### Phrase 번역 초기화!
                UntranslatePhrase(LVST.FocusedItem.Index)
            End If
        End If
    End Sub

    Private Sub WBTT_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs) Handles WBTT.Navigated
        WBTT.Document.Body.ScrollLeft = g_WBScrollLeft
        WBTT.Document.Body.ScrollTop = g_WBScrollTop
    End Sub
End Class

