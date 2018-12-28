Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class MainFrm
    Public Const MAX_HTMLSEGMENT As Integer = 1000
    Public Const MAX_HTMLSUBSEGMENT As Integer = 3000
    Public Const MAX_PHRASE As Integer = 5000

    Dim g_objWord As Word.Application
    Dim g_objDoc As Word.Document

    Dim g_FontSentence As New Font("맑은 고딕", 10, FontStyle.Regular)
    Dim g_nPrevLVSTIndex As Integer = 0

    Dim g_sHTMLOrg As String
    Dim g_sHTMLHEAD As String
    Dim g_sHTMLBODY As String
    Dim g_sHTMLTAIL As String

    Dim g_sHTMLTrs As String

    Public Structure HTML_SEG
        Dim Content As String
        Dim SubSegCount As Integer
    End Structure

    Public Structure HTML_SUBSEG
        Dim Content As String
        Dim IndexInSeg As Integer
        Dim ParentSegID As Integer
        Dim PhraseCount As Integer
    End Structure

    Public Structure HTML_PHRASE
        Dim Content As String
        Dim IndexInSubSeg As Integer
        Dim ParentSubSegID As Integer
        Dim ParentSegID As Integer
    End Structure

    Public Structure HTML_NODE_INFO
        Dim NodeName As String
        Dim ElementCount As Integer
        Dim ElementName() As String
    End Structure

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
    Dim g_sFN_DOCX_T As String = Nothing

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
            If Mid(g_sHTMLOrg, i, 4) = "<div" Then
                For j = i + 4 To g_sHTMLOrg.Length
                    If Mid(g_sHTMLOrg, j, 1) = ">" Then
                        g_sHTMLHEAD = Strings.Left(g_sHTMLOrg, j) & vbCrLf
                        g_sHTMLBODY = Mid(g_sHTMLBODY, j + 1, g_sHTMLOrg.Length)
                        Exit For
                    End If
                Next
            End If
        Next

        For i = 1 To g_sHTMLOrg.Length
            If Mid(g_sHTMLOrg, i, 6) = "</div>" Then
                g_sHTMLTAIL = vbCrLf & Mid(g_sHTMLOrg, i, g_sHTMLOrg.Length)
                g_sHTMLBODY = Strings.Left(g_sHTMLBODY, i - 1)
                Exit For
            End If
        Next

        '### 이제부터가 파싱 시작! ###
        g_nHTMLSegsCount = 0
        For i = 1 To g_sHTMLBODY.Length
            If Mid(g_sHTMLBODY, i, 2) = "<p" Then
                g_nHTMLSegsCount = g_nHTMLSegsCount + 1
                For j = i + 1 To g_sHTMLBODY.Length
                    If Mid(g_sHTMLBODY, j, 4) = "</p>" Then
                        g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content = Mid(g_sHTMLBODY, i, j - i + 4)
                        g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content = g_HTMLOrgSegs(g_nHTMLSegsCount - 1).Content.Replace(vbCrLf, " ")
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

    Public Function HTML_GetNodeInfo(ByVal Node As String) As HTML_NODE_INFO
        HTML_GetNodeInfo.NodeName = "NOTHING"
        HTML_GetNodeInfo.ElementCount = 0
        ReDim HTML_GetNodeInfo.ElementName(0)

        If Node.Length = 0 Then Exit Function


        If Strings.Left(Node, 1) = "<" Then Node = Strings.Right(Node, Node.Length - 1)

        Dim T_Node_Info As HTML_NODE_INFO

        Dim SpaceCount As Integer = 0
        For i = 1 To Node.Length
            If Mid(Node, i, 1) = " " Then
                SpaceCount = SpaceCount + 1
            End If
        Next

        If SpaceCount = 0 Then
            '### <b>처럼 Element가 없는 Node일 때!
            T_Node_Info.NodeName = Strings.Left(Node, Node.Length - 1)
            T_Node_Info.ElementCount = 0
            ReDim T_Node_Info.ElementName(0)
            T_Node_Info.ElementName(0) = Nothing
            Exit Function
        End If

        T_Node_Info.NodeName = Nothing
        T_Node_Info.ElementCount = 0
        ReDim T_Node_Info.ElementName(SpaceCount - 1)
        For i = 0 To SpaceCount - 1
            T_Node_Info.ElementName(i) = Nothing
        Next

        Dim ElementCount As Integer = 0
        Dim Prev_i As Integer = 0

        For i = 1 To Node.Length
            If Mid(Node, i, 1) = " " Or Mid(Node, i, 1) = ">" Then
                ElementCount = ElementCount + 1

                Select Case ElementCount
                    Case 1
                        T_Node_Info.NodeName = LCase(Mid(Node, 1, i - 1))
                        Prev_i = i + 1
                    Case Else
                        For j = Prev_i To Node.Length
                            If Mid(Node, j, 1) = "=" Then
                                T_Node_Info.ElementCount = T_Node_Info.ElementCount + 1
                                T_Node_Info.ElementName(ElementCount - 2) = LCase(Mid(Node, Prev_i, j - Prev_i))
                                Exit For
                            End If
                        Next
                        Prev_i = i + 1
                End Select
            End If
        Next

        HTML_GetNodeInfo = T_Node_Info
    End Function

    Private Function HTML_Encode_Entities(ByVal Value As String) As String
        HTML_Encode_Entities = Value

        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#192;", "&Agrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#193;", "&Aacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#194;", "&Acirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#195;", "&Atilde;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#196;", "&Auml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#197;", "&Aring;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#198;", "&AElig;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#199;", "&Ccedil;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#200;", "&Egrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#201;", "&Eacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#202;", "&Ecirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#203;", "&Euml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#204;", "&Igrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#205;", "&Iacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#206;", "&Icirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#207;", "&Iuml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#208;", "&ETH;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#209;", "&Ntilde;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#210;", "&Ograve;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#211;", "&Oacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#212;", "&Ocirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#213;", "&Otilde;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#214;", "&Ouml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#216;", "&Oslash;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#217;", "&Ugrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#218;", "&Uacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#219;", "&Ucirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#220;", "&Uuml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#221;", "&Yacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#222;", "&THORN;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#223;", "&szlig;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#224;", "&agrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#225;", "&aacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#226;", "&acirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#227;", "&atilde;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#228;", "&auml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#229;", "&aring;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#230;", "&aelig;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#231;", "&ccedil;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#232;", "&egrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#233;", "&eacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#234;", "&ecirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#235;", "&euml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#236;", "&igrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#237;", "&iacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#238;", "&icirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#239;", "&iuml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#240;", "&eth;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#241;", "&ntilde;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#242;", "&ograve;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#243;", "&oacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#244;", "&ocirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#245;", "&otilde;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#246;", "&ouml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#248;", "&oslash;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#249;", "&ugrave;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#250;", "&uacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#251;", "&ucirc;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#252;", "&uuml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#253;", "&yacute;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#254;", "&thorn;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#255;", "&yuml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#338;", "&OElig;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#339;", "&oelig;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#352;", "&Scaron;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#353;", "&scaron;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#376;", "&Yuml;")
        HTML_Encode_Entities = Replace(HTML_Encode_Entities, "&#402;", "&fnof;")
    End Function

    Private Function HTML_Encode_Space(ByVal Value As String) As String
        Dim SpaceStartPos As Long = 0
        Dim SpaceCount As Long = 0

        For i = 1 To Value.Length
            If Mid(Value, i, 1) = " " Then
                If SpaceCount = 0 Then SpaceStartPos = i
                SpaceCount = SpaceCount + 1
            Else
                If SpaceCount >= 2 Then
                    Dim HTMLSpaceString As String = Nothing

                    For j = 1 To SpaceCount - 1
                        HTMLSpaceString = HTMLSpaceString & "&nbsp;"
                    Next
                    HTMLSpaceString = HTMLSpaceString & " "
                    Value = String_ReplaceByPos(Value, SpaceStartPos, i - 1, HTMLSpaceString)
                End If
                SpaceCount = 0
            End If
        Next
        HTML_Encode_Space = Value
    End Function

    Public Function HTML_Encode(ByVal Value As String) As String
        Dim T_Output As String

        T_Output = Net.WebUtility.HtmlEncode(Value)
        T_Output = HTML_Encode_Entities(T_Output)
        T_Output = HTML_Encode_Space(T_Output)

        HTML_Encode = T_Output
    End Function

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

    Public Sub UpdatePhrase(ByVal CurItemID As Integer)
        Dim CurSegID As Integer = g_STPhrases(CurItemID).ParentSegID
        Dim CurSubSegRelID As Integer = g_HTMLSubSegs(g_STPhrases(CurItemID).ParentSubSegID).IndexInSeg
        Dim CurPhraseRelID As Integer = g_STPhrases(CurItemID).IndexInSubSeg

        g_sTTPhrases(CurItemID) = TBTTSentence.Text

        Dim bPhraseDivided As Boolean = False

        If g_HTMLSubSegs(g_STPhrases(CurItemID).ParentSubSegID).PhraseCount > 1 Then
            '### 문장 구분이 있으면
            bPhraseDivided = True
        End If

        '### TT에 아무 내용도 없으면? 번역하지 말기!
        If g_sTTPhrases(CurItemID) = "" Then Exit Sub

        Dim sTrgEncoded As String = HTML_Encode(g_sTTPhrases(CurItemID))

        UpdateSegment(CurSegID, CurSubSegRelID, CurPhraseRelID, bPhraseDivided, sTrgEncoded)


        '### 이제 HTML 수정!!
        g_sHTMLTrs = g_sHTMLHEAD
        For i = 0 To g_nHTMLSegsCount - 1
            If g_bSegTranslated(i) = True Then
                g_sHTMLTrs = g_sHTMLTrs & g_sHTMLTrsSegs(i)
            Else
                g_sHTMLTrs = g_sHTMLTrs & g_HTMLOrgSegs(i).Content
            End If
        Next

        g_sHTMLTrs = g_sHTMLTrs & g_sHTMLTAIL

        g_bPhraseTranslated(CurItemID) = True
    End Sub

    Private Sub InitiateVariables()
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
        LVST.Columns(1).Width = 40
        LVST.Columns(1).TextAlign = HorizontalAlignment.Center
        LVST.Columns(2).Width = 300
        LVST.Columns(3).Width = 280

        InitiateVariables()
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

        g_objWord = New Word.Application
        g_objDoc = New Word.Document

        g_objDoc = g_objWord.Documents.Add(g_sFN_HTML)
        g_objDoc.SaveAs2(g_sFN_DOCX_T, FileFormat:=16) '### 16 = DOCX format ###

        g_objDoc.Close()
        g_objWord.Quit()

        g_objDoc = Nothing
        g_objWord = Nothing

    End Sub

    Private Sub LVST_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LVST.SelectedIndexChanged
        If LVST.FocusedItem IsNot Nothing Then
            TBSTSentence.Text = g_STPhrases(LVST.FocusedItem.Index).Content
            TBTTSentence.Text = g_sTTPhrases(LVST.FocusedItem.Index)
        End If
    End Sub

    Private Sub LVST_KeyDown(sender As Object, e As KeyEventArgs) Handles LVST.KeyDown
        If e.KeyCode = Keys.F1 Or e.KeyCode = Keys.Enter Then
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
                '### 달라진 게 있으면 반영!
                UpdatePhrase(g_nPrevLVSTIndex)
                LVST.Items(g_nPrevLVSTIndex).SubItems(3).Text = g_sTTPhrases(g_nPrevLVSTIndex)

                '### 번역한 HTML을 저장한다
                FileIO.FileSystem.WriteAllText(g_sFN_HTML, g_sHTMLTrs, False)
                WBTT.Navigate(g_sFN_HTML)

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

    Private Sub TBTTSentence_TextChanged(sender As Object, e As EventArgs) Handles TBTTSentence.TextChanged

    End Sub
End Class

