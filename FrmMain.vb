Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class MainFrm
    Public Const MAX_HTMLSEGMENT As Integer = 10000
    Public Const MAX_HTMLSUBSEGMENT As Integer = 30000
    Public Const MAX_PHRASE As Integer = 50000

    Dim g_objWord As Word.Application
    Dim g_objDoc As Word.Document

    Dim g_FontSentence As New Font("맑은 고딕", 10, FontStyle.Regular)
    Dim g_nPrevLVSTIndex As Integer = 0

    Dim g_sHTMLOrg As String
    Dim g_sHTMLHEAD As String
    Dim g_sHTMLBODY As String
    Dim g_sHTMLTAIL As String

    Dim g_sHTMLTrs As String

    Dim g_HTMLOrgSegs(MAX_HTMLSEGMENT) As HTML_SEG
    Dim g_sHTMLTrsSegs(MAX_HTMLSEGMENT) As String
    Dim g_nHTMLSegsCount As Integer = 0

    Dim g_HTMLSubSegs(MAX_HTMLSUBSEGMENT) As HTML_SUBSEG
    Dim g_nHTMLSubSegsCount As Integer = 0

    Dim g_STPhrases(MAX_PHRASE) As HTML_PHRASE
    Dim g_nSTPhrasesCount As Integer = 0
    Dim g_sTTPhrases(MAX_PHRASE) As String

    Dim g_bPhraseTranslated(MAX_PHRASE) As Boolean
    Dim g_bSegTranslated(MAX_HTMLSEGMENT) As Boolean

    Dim g_sDIR_HTML As String = Nothing
    Dim g_sFN_HTML As String = Nothing

    Dim g_WBScrollTop As Integer
    Dim g_WBScrollLeft As Integer

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

    Public Sub HTML_ParseIntoPhrases()

        If FileIO.FileSystem.FileExists(g_sFN_HTML) = False Then Exit Sub

        g_sHTMLOrg = FileIO.FileSystem.ReadAllText(g_sFN_HTML, System.Text.Encoding.Default)
        g_sHTMLBODY = g_sHTMLOrg
        g_sHTMLTrs = g_sHTMLOrg

        For i = 1 To g_sHTMLOrg.Length
            If Mid(g_sHTMLOrg, i, 5) = "<body" Then
                For j = i + 5 To g_sHTMLOrg.Length
                    If Mid(g_sHTMLOrg, j, 1) = ">" Then
                        g_sHTMLHEAD = Strings.Left(g_sHTMLOrg, j) & vbCrLf
                        g_sHTMLBODY = Mid(g_sHTMLBODY, j + 1, g_sHTMLOrg.Length)
                        Exit For
                    End If
                Next
            End If
        Next

        For i = 1 To g_sHTMLOrg.Length
            If Mid(g_sHTMLOrg, i, 7) = "</body>" Then
                g_sHTMLTAIL = vbCrLf & Mid(g_sHTMLOrg, i, g_sHTMLOrg.Length)
                g_sHTMLBODY = Strings.Left(g_sHTMLBODY, i - 1)
                Exit For
            End If
        Next

        '### 이제부터가 파싱 시작! ###
        g_nHTMLSegsCount = 0
        For i = 1 To g_sHTMLBODY.Length
            If Mid(g_sHTMLBODY, i, 2) = "<p" Then
                '### <p ~>인 경우!
                g_nHTMLSegsCount = g_nHTMLSegsCount + 1
                For j = i + 1 To g_sHTMLBODY.Length
                    If Mid(g_sHTMLBODY, j, 4) = "</p>" Then
                        g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content = Mid(g_sHTMLBODY, i, j - i + 4)
                        g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content = g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content.Replace(vbCrLf, " ")
                        i = j + 3
                        Exit For
                    End If
                Next
            ElseIf Mid(g_sHTMLBODY, i, 1) = "<" Then
                '### <p ~>는 아니지만 노드가 있는 경우 다음 <p ~> / </body> 전까지 저장해 두자 (그래야 나중에 복원할 수 있으니까)

                If Mid(g_sHTMLBODY, i, 7) = "</body>" Then
                    '### </body>나오면 Phrase 파싱 끝내기!
                    Exit For
                End If

                g_nHTMLSegsCount = g_nHTMLSegsCount + 1
                For j = i + 1 To g_sHTMLBODY.Length
                    If Mid(g_sHTMLBODY, j, 2) = "<p" Or Mid(g_sHTMLBODY, j, 7) = "</body>" Then
                        g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content = Mid(g_sHTMLBODY, i, j - i)
                        g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content = g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content.Replace(vbCrLf, " ")
                        i = j - 1
                        Exit For
                    End If
                Next
            End If
        Next

        If g_nHTMLSegsCount > 0 Then
            Dim HTMLNode As HTML_NODE_INFO

            For i = 0 To g_nHTMLSegsCount - 1
                For j = 1 To g_HTMLOrgSegs(i).Content.Length
                    If Mid(g_HTMLOrgSegs(i).Content, j, 1) = "<" Then
                        For k = j + 1 To g_HTMLOrgSegs(i).Content.Length
                            If Mid(g_HTMLOrgSegs(i).Content, k, 1) = ">" Then
                                '### Node 하나를 찾음
                                HTMLNode = HTML_GetNodeInfo(Mid(g_HTMLOrgSegs(i).Content, j, k - j + 1))

                                If HTMLNode.NodeName = "span" And HTMLNode.ElementCount = 1 And HTMLNode.ElementName(0) = "lang" Then
                                    '### 찾은 Node가 <span lang=~~>일 경우 없애자!!★★
                                    For l = k + 1 To g_HTMLOrgSegs(i).Content.Length
                                        If Mid(g_HTMLOrgSegs(i).Content, l, 7) = "</span>" Then
                                            Dim sLeft As String = Nothing
                                            Dim sMid As String = Nothing
                                            Dim sRight As String = Nothing

                                            sLeft = Strings.Left(g_HTMLOrgSegs(i).Content, j - 1)
                                            sMid = Mid(g_HTMLOrgSegs(i).Content, k + 1, l - k - 1)
                                            sRight = Strings.Mid(g_HTMLOrgSegs(i).Content, l + 7, g_HTMLOrgSegs(i).Content.Length)

                                            g_HTMLOrgSegs(i).Content = sLeft & sMid & sRight
                                            Exit For
                                        End If
                                    Next

                                    Exit For
                                Else
                                    '### 그렇지 않은 Node일 경우
                                    j = k
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
            Next

            Dim nSubSegCountPerSeg As Integer = 0
            For i = 0 To g_nHTMLSegsCount - 1
                nSubSegCountPerSeg = 0

                If Strings.Left(g_HTMLOrgSegs(i).Content, 2) = "<p" Then
                    '### <p ~>일 때만 SubSeg를 찾는다!★★
                    For j = 1 To g_HTMLOrgSegs(i).Content.Length
                        If Mid(g_HTMLOrgSegs(i).Content, j, 1) = "<" Then
                            For k = j + 1 To g_HTMLOrgSegs(i).Content.Length
                                If Mid(g_HTMLOrgSegs(i).Content, k, 1) = ">" Then
                                    '### Node 하나를 찾음
                                    j = k
                                    Exit For
                                End If
                            Next
                        Else
                            For k = j To g_HTMLOrgSegs(i).Content.Length
                                If Mid(g_HTMLOrgSegs(i).Content, k, 1) = "<" Then
                                    '### Node 사이의 문자열을 찾은 경우
                                    nSubSegCountPerSeg = nSubSegCountPerSeg + 1
                                    g_nHTMLSubSegsCount = g_nHTMLSubSegsCount + 1
                                    g_HTMLSubSegs(g_nHTMLSubSegsCount - 1).Content = Mid(g_HTMLOrgSegs(i).Content, j, k - j)
                                    g_HTMLSubSegs(g_nHTMLSubSegsCount - 1).IndexInSeg = nSubSegCountPerSeg - 1
                                    g_HTMLSubSegs(g_nHTMLSubSegsCount - 1).ParentSegID = i

                                    j = k - 1
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                End If

                g_HTMLOrgSegs(i).SubSegCount = nSubSegCountPerSeg
            Next


            '### 문장 구분 ###
            Dim nPhraseCountPerSubSeg As Integer = 0
            g_nSTPhrasesCount = 0

            For i = 0 To g_nHTMLSubSegsCount - 1
                nPhraseCountPerSubSeg = 0

                For j = 1 To g_HTMLSubSegs(i).Content.Length
                    If Mid(g_HTMLSubSegs(i).Content, j, 1) = "." Then

                        If j = 1 Then
                            '### 맨 처음인데 .이 찍혀있으면 나누지 않음
                        ElseIf Mid(g_HTMLSubSegs(i).Content, j - 1, 1) = ">" Or Mid(g_HTMLSubSegs(i).Content, j - 1, 1) = "." Then
                            '### 맨 처음 글자가 .이거나 바로 앞 글자도 .이었을 경우 나누지 않음
                        Else

                            If j = g_HTMLSubSegs(i).Content.Length Then
                                '### 마지막이 .이면 당연히 나눈다!
                                g_nSTPhrasesCount = g_nSTPhrasesCount + 1
                                nPhraseCountPerSubSeg = nPhraseCountPerSubSeg + 1
                                g_STPhrases(g_nSTPhrasesCount - 1).Content = Mid(g_HTMLSubSegs(i).Content, 1, j)
                                g_STPhrases(g_nSTPhrasesCount - 1).IndexInSubSeg = nPhraseCountPerSubSeg - 1
                                g_STPhrases(g_nSTPhrasesCount - 1).ParentSubSegID = i
                                g_STPhrases(g_nSTPhrasesCount - 1).ParentSegID = g_HTMLSubSegs(g_STPhrases(g_nSTPhrasesCount - 1).ParentSubSegID).ParentSegID
                                g_HTMLSubSegs(i).Content = Trim(Mid(g_HTMLSubSegs(i).Content, j + 1, g_HTMLSubSegs(i).Content.Length))
                                j = 0

                            ElseIf Mid(g_HTMLSubSegs(i).Content, j + 1, 1) = " " Then
                                '### 다음 글자가 띄어쓰기여도 나눈다!! ★★ (. 다음이 띄어쓰기가 아니면 안 나눈다!!★)
                                g_nSTPhrasesCount = g_nSTPhrasesCount + 1
                                nPhraseCountPerSubSeg = nPhraseCountPerSubSeg + 1
                                g_STPhrases(g_nSTPhrasesCount - 1).Content = Mid(g_HTMLSubSegs(i).Content, 1, j)
                                g_STPhrases(g_nSTPhrasesCount - 1).IndexInSubSeg = nPhraseCountPerSubSeg - 1
                                g_STPhrases(g_nSTPhrasesCount - 1).ParentSubSegID = i
                                g_STPhrases(g_nSTPhrasesCount - 1).ParentSegID = g_HTMLSubSegs(g_STPhrases(g_nSTPhrasesCount - 1).ParentSubSegID).ParentSegID
                                g_HTMLSubSegs(i).Content = Trim(Mid(g_HTMLSubSegs(i).Content, j + 1, g_HTMLSubSegs(i).Content.Length))
                                j = 0
                            End If
                        End If
                    End If
                Next

                Dim T_Content As String = Nothing
                T_Content = g_HTMLSubSegs(i).Content.Replace("&nbsp;", "") '### &nbsp; 는 빈 내용이므로 번역이 필요 없으니까!!

                If T_Content = " " Then T_Content = "" '### 띄어쓰기 하나만 남았으면 없애자!

                If T_Content <> "" Then '### 아직 내용이 남아 있으면 새로 추가!!
                    g_nSTPhrasesCount = g_nSTPhrasesCount + 1
                    nPhraseCountPerSubSeg = nPhraseCountPerSubSeg + 1
                    g_STPhrases(g_nSTPhrasesCount - 1).Content = Trim(g_HTMLSubSegs(i).Content)
                    g_STPhrases(g_nSTPhrasesCount - 1).IndexInSubSeg = nPhraseCountPerSubSeg - 1
                    g_STPhrases(g_nSTPhrasesCount - 1).ParentSubSegID = i
                    g_STPhrases(g_nSTPhrasesCount - 1).ParentSegID = g_HTMLSubSegs(g_STPhrases(g_nSTPhrasesCount - 1).ParentSubSegID).ParentSegID
                End If

                g_HTMLSubSegs(i).PhraseCount = nPhraseCountPerSubSeg
            Next

        End If

        LVST.Items.Clear()
        If g_nSTPhrasesCount > 0 Then
            For i = 0 To g_nSTPhrasesCount - 1
                g_STPhrases(i).Content = Net.WebUtility.HtmlDecode(g_STPhrases(i).Content)
                'LVST.Items.Add(New ListViewItem({"", i + 1, g_STPhrases(i).Content, g_STPhrases(i).ParentSegID})) '### 디버그용
                LVST.Items.Add(New ListViewItem({"", i + 1, g_STPhrases(i).Content, ""}))
            Next
        End If

    End Sub

    Private Function UpdateSubSegment(ByVal PhraseID As Integer, ByVal Content As String, ByVal SubSeg As String) As String
        UpdateSubSegment = Nothing

        '### 변수 선언 및 초기화
        Dim nPhraseCount As Integer = 0
        Dim sTempPhrases(MAX_PHRASE) As String

        For i = 0 To MAX_PHRASE
            sTempPhrases(MAX_PHRASE) = Nothing
        Next

        For j = 1 To SubSeg.Length
            If Mid(SubSeg, j, 1) = "." Then

                If j = 1 Then
                    '### 맨 처음인데 .이 찍혀있으면 나누지 않음
                ElseIf Mid(SubSeg, j - 1, 1) = ">" Or Mid(SubSeg, j - 1, 1) = "." Then
                    '### 맨 처음 글자가 .이거나 바로 앞 글자도 .이었을 경우 나누지 않음
                Else

                    If j = SubSeg.Length Then
                        '### 마지막이 .이면 당연히 나눈다!
                        nPhraseCount = nPhraseCount + 1
                        sTempPhrases(nPhraseCount - 1) = Mid(SubSeg, 1, j)
                        SubSeg = Mid(SubSeg, j + 1, SubSeg.Length)
                        j = 0

                    ElseIf Mid(SubSeg, j + 1, 1) = " " Then
                        '### 다음 글자가 띄어쓰기여도 나눈다!! ★★ (. 다음이 띄어쓰기가 아니면 안 나눈다!!★)
                        nPhraseCount = nPhraseCount + 1
                        sTempPhrases(nPhraseCount - 1) = Mid(SubSeg, 1, j)
                        SubSeg = Mid(SubSeg, j + 1, SubSeg.Length)
                        j = 0
                    End If
                End If
            End If
        Next

        Dim TempSubSeg As String = SubSeg.Replace("&nbsp;", "") '### &nbsp; 는 빈 내용이므로 번역이 필요 없으니까!!
        If TempSubSeg = " " Then TempSubSeg = "" '### 띄어쓰기 하나만 남았으면 없애자!

        If TempSubSeg <> "" Then '### 아직 내용이 남아 있으면? Phrase 하나 추가!!
            nPhraseCount = nPhraseCount + 1
            sTempPhrases(nPhraseCount - 1) = SubSeg
        End If

        sTempPhrases(PhraseID) = Content

        If PhraseID = nPhraseCount Then
            '### 마지막 Phrase면 마지막에 점이 있든 없든 상관이 없다!
        Else
            '### 마지막 Phrase가 아니라면 반드시 '. '가 있어야 한다!!!
            If Strings.Right(sTempPhrases(PhraseID), 1) <> " " Then
                If Strings.Right(sTempPhrases(PhraseID), 1) <> "." Then sTempPhrases(PhraseID) = sTempPhrases(PhraseID) & "."
                sTempPhrases(PhraseID) = sTempPhrases(PhraseID) & " "
            Else
                Dim FountPeriod As Boolean = False

                For i = 1 To sTempPhrases(PhraseID).Length
                    If Mid(sTempPhrases(PhraseID), i, 2) = ". " Then
                        FountPeriod = True
                    End If
                Next

                If FountPeriod = False Then
                    sTempPhrases(PhraseID) = sTempPhrases(PhraseID) & "."
                End If
            End If
        End If


        For i = 0 To nPhraseCount - 1
            UpdateSubSegment = UpdateSubSegment & sTempPhrases(i) & " "
        Next

        UpdateSubSegment = UpdateSubSegment
    End Function

    Private Sub UpdateSegment(ByVal SegID As Integer, ByVal SubSegID As Integer, ByVal PhraseID As Integer,
                              ByVal PhraseDivided As Boolean, ByVal Content As String)

        Dim sSrcSegment As String = Nothing

        Select Case g_bSegTranslated(SegID)
            Case True
                sSrcSegment = g_sHTMLTrsSegs(SegID)
            Case False
                sSrcSegment = g_HTMLOrgSegs(SegID).Content
        End Select


        Dim sSrcLeft As String = Nothing
        Dim sSrcRight As String = Nothing

        Dim nSubSegCount As Integer = 0

        For i = 1 To sSrcSegment.Length
            If Mid(sSrcSegment, i, 1) = "<" Then
                For j = i + 1 To sSrcSegment.Length
                    If Mid(sSrcSegment, j, 1) = ">" Then
                        i = j
                        Exit For
                    End If
                Next
            Else
                For j = i To sSrcSegment.Length
                    If Mid(sSrcSegment, j, 1) = "<" Then

                        If nSubSegCount = SubSegID Then
                            '### SubSegment 찾음!! 번역하자 ###

                            Select Case PhraseDivided
                                Case True
                                    '### Phrase 구분이 되어 있으면?
                                    Dim sSubSeg As String = Mid(sSrcSegment, i, j - i)
                                    Dim UpdatedContent As String = UpdateSubSegment(PhraseID, Content, sSubSeg)

                                    sSrcLeft = Strings.Left(sSrcSegment, i - 1)
                                    sSrcRight = Mid(sSrcSegment, j, sSrcSegment.Length)

                                    g_sHTMLTrsSegs(SegID) = sSrcLeft & UpdatedContent & sSrcRight
                                    g_bSegTranslated(SegID) = True

                                Case False
                                    '### Phrase 구분이 없으면?
                                    sSrcLeft = Strings.Left(sSrcSegment, i - 1)
                                    sSrcRight = Mid(sSrcSegment, j, sSrcSegment.Length)

                                    g_sHTMLTrsSegs(SegID) = sSrcLeft & Content & sSrcRight
                                    g_bSegTranslated(SegID) = True
                            End Select

                            Exit Sub
                        End If

                        i = j - 1
                        nSubSegCount = nSubSegCount + 1

                        Exit For
                    End If
                Next
            End If
        Next

    End Sub

    Public Sub TranslatePhrase(ByVal CurItemID As Integer)
        Dim CurSegID As Integer = g_STPhrases(CurItemID).ParentSegID
        Dim CurSubSegRelID As Integer = g_HTMLSubSegs(g_STPhrases(CurItemID).ParentSubSegID).IndexInSeg
        Dim CurPhraseRelID As Integer = g_STPhrases(CurItemID).IndexInSubSeg

        g_sTTPhrases(CurItemID) = TBTTSentence.Text

        Dim bPhraseDivided As Boolean = False

        If g_HTMLSubSegs(g_STPhrases(CurItemID).ParentSubSegID).PhraseCount > 1 Then
            '### 문장 구분이 있다!
            bPhraseDivided = True
        End If

        '### TT에 아무 내용도 없으면? 번역하지 말기!
        If g_sTTPhrases(CurItemID) = "" Then Exit Sub

        Dim sTrgEncoded As String = HTML_Encode(g_sTTPhrases(CurItemID))
        UpdateSegment(CurSegID, CurSubSegRelID, CurPhraseRelID, bPhraseDivided, sTrgEncoded)

        '### ListView에 수정!
        LVST.Items(CurItemID).SubItems(3).Text = g_sTTPhrases(CurItemID)

        '### 이제 HTML 수정!!
        UpdateHTMLCode()

        g_bPhraseTranslated(CurItemID) = True
    End Sub

    Public Sub UpdateHTMLCode()
        g_sHTMLTrs = g_sHTMLHEAD
        For i = 0 To g_nHTMLSegsCount - 1
            If g_bSegTranslated(i) = True Then
                g_sHTMLTrs = g_sHTMLTrs & g_sHTMLTrsSegs(i)
            Else
                g_sHTMLTrs = g_sHTMLTrs & g_HTMLOrgSegs(i).Content
            End If
        Next

        g_sHTMLTrs = g_sHTMLTrs & g_sHTMLTAIL

        '### 바뀐 HTML을 저장한다
        g_WBScrollTop = WBTT.Document.Body.ScrollTop
        g_WBScrollLeft = WBTT.Document.Body.ScrollLeft

        FileIO.FileSystem.WriteAllText(g_sFN_HTML, g_sHTMLTrs, False)
        WBTT.Navigate(g_sFN_HTML)
    End Sub

    Public Sub UntranslatePhrase(ByVal PhraseID As Integer)
        TBTTSentence.Text = g_STPhrases(PhraseID).Content
        TranslatePhrase(PhraseID)

        TBTTSentence.Text = Nothing
        g_sTTPhrases(PhraseID) = Nothing
        LVST.Items(PhraseID).SubItems(3).Text = Nothing
        g_bPhraseTranslated(PhraseID) = False
    End Sub

    Private Sub InitializeVariables()
        '### 변수 초기화★ ###

        g_sHTMLOrg = Nothing
        g_sHTMLHEAD = Nothing
        g_sHTMLBODY = Nothing
        g_sHTMLTAIL = Nothing
        g_sHTMLTrs = Nothing

        g_nHTMLSegsCount = 0
        For i = 0 To MAX_HTMLSEGMENT
            g_HTMLOrgSegs(i).Content = Nothing
            g_HTMLOrgSegs(i).SubSegCount = 0
            g_bSegTranslated(i) = False
        Next

        g_nHTMLSubSegsCount = 0
        For i = 0 To MAX_HTMLSUBSEGMENT
            g_HTMLSubSegs(i).Content = Nothing
            g_HTMLSubSegs(i).IndexInSeg = 0
            g_HTMLSubSegs(i).ParentSegID = 0
            g_HTMLSubSegs(i).PhraseCount = 0
        Next

        g_nSTPhrasesCount = 0
        For i = 0 To MAX_PHRASE
            g_STPhrases(i).Content = Nothing
            g_STPhrases(i).ParentSegID = 0
            g_STPhrases(i).ParentSubSegID = 0

            g_sTTPhrases(i) = Nothing
            g_bPhraseTranslated(i) = False
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

        End

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
            TBSTSentence.Text = g_STPhrases(LVST.FocusedItem.Index).Content
            TBTTSentence.Text = g_sTTPhrases(LVST.FocusedItem.Index)
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

            If g_sTTPhrases(g_nPrevLVSTIndex) <> TBTTSentence.Text Then
                '### 달라진 게 있으면 반영하자!

                If TBTTSentence.Text = "" Then
                    '### TBTTSentence.Text가 Null인 경우
                    TBTTSentence.Text = g_sTTPhrases(g_nPrevLVSTIndex)
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

