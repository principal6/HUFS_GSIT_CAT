Module ModHTML

    Public Const MAX_HTMLSEGMENT As Integer = 10000
    Public Const MAX_HTMLSUBSEGMENT As Integer = 30000
    Public Const MAX_PHRASE As Integer = 50000
    Public Const PHRASE_PER_SUBSEG As Integer = 1000

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

    Public Structure HTML_NODEANDTEXT
        Dim NODEANDTEXT As String
        Dim StartIndex As Integer
        Dim EndIndex As Integer
    End Structure

    Public Function HTML_DeleteNodeHeadTail(ByVal HTMLText As String, ByVal NodeName As String) As String
        HTML_DeleteNodeHeadTail = Nothing

        If HTMLText = Nothing Then Exit Function

        Dim sNodeHead As String = Nothing
        Dim sNodeTail As String = Nothing


        For i = 1 To HTMLText.Length
            If Mid(HTMLText, i, 1 + NodeName.Length) = "<" & NodeName Then
                For j = i + NodeName.Length To HTMLText.Length
                    If Mid(HTMLText, j, 1) = ">" Then
                        '### Found the Node Head
                        sNodeHead = Mid(HTMLText, i, j - i + 1)
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next

        '### NodeHead가 없으면 종료!★
        If sNodeHead = Nothing Then Exit Function

        sNodeTail = "</" & NodeName & ">"

        HTMLText = HTMLText.Replace(sNodeHead, "")
        HTMLText = HTMLText.Replace(sNodeTail, "")

        HTML_DeleteNodeHeadTail = HTMLText
    End Function

    Public Function HTML_GetFirstNodeHead(ByVal HTMLText As String) As String
        HTML_GetFirstNodeHead = Nothing

        If HTMLText = Nothing Then Exit Function

        For i = 1 To HTMLText.Length
            If Mid(HTMLText, i, 1) = "<" Then
                For j = i + 1 To HTMLText.Length
                    If Mid(HTMLText, j, 1) = ">" Then
                        '### 찾음
                        HTML_GetFirstNodeHead = Mid(HTMLText, i, j - i + 1)
                        Exit For

                    End If
                Next
                Exit For
            End If
        Next
    End Function

    Public Function HTML_GetNodeHeadEndPos(ByVal HTMLText As String, ByVal NodeName As String) As Integer
        HTML_GetNodeHeadEndPos = 0

        If HTMLText = Nothing Then Exit Function

        For i = 1 To HTMLText.Length
            If Mid(HTMLText, i, NodeName.Length + 1) = "<" & NodeName Then
                For j = i + 5 To HTMLText.Length
                    If Mid(HTMLText, j, 1) = ">" Then
                        HTML_GetNodeHeadEndPos = j

                        Exit For
                    End If
                Next

                Exit For
            End If
        Next
    End Function

    Public Function HTML_GetBeforeNodeHead(ByVal HTMLText As String, ByVal NodeName As String) As String
        HTML_GetBeforeNodeHead = Nothing

        If HTMLText = Nothing Then Exit Function

        HTML_GetBeforeNodeHead = Strings.Left(HTMLText, HTML_GetNodeHeadEndPos(HTMLText, NodeName)) & vbCrLf
    End Function

    Public Function HTML_GetNodeTailStartPos(ByVal HTMLText As String, ByVal NodeName As String) As Integer
        HTML_GetNodeTailStartPos = 0

        If HTMLText = Nothing Then Exit Function

        For i = 1 To HTMLText.Length
            If Mid(HTMLText, i, NodeName.Length + 3) = "</" & NodeName & ">" Then
                HTML_GetNodeTailStartPos = i
                Exit For
            End If
        Next
    End Function

    Public Function HTML_GetAfterNodeTail(ByVal HTMLText As String, ByVal NodeName As String) As String
        HTML_GetAfterNodeTail = Nothing

        If HTMLText = Nothing Then Exit Function

        HTML_GetAfterNodeTail = vbCrLf & Mid(HTMLText, HTML_GetNodeTailStartPos(HTMLText, NodeName), HTMLText.Length)
    End Function

    Public Function HTML_GetIntraNode(ByVal HTMLText As String, ByVal NodeName As String) As String
        HTML_GetIntraNode = Nothing

        If HTMLText = Nothing Then Exit Function

        Dim T_StartPos As Integer = 0
        Dim T_EndPos As Integer = 0

        For i = 1 To HTMLText.Length
            If Mid(HTMLText, i, NodeName.Length + 1) = "<" & NodeName Then
                For j = i + 5 To HTMLText.Length
                    If Mid(HTMLText, j, 1) = ">" Then
                        T_StartPos = j + 1
                        i = T_StartPos
                        Exit For
                    End If
                Next
            ElseIf Mid(HTMLText, i, NodeName.Length + 3) = "</" & NodeName & ">" Then
                T_EndPos = i
                Exit For
            End If
        Next

        HTML_GetIntraNode = Mid(HTMLText, T_StartPos, T_EndPos - T_StartPos)
    End Function

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
        Dim SpaceStartPos As Integer = 0
        Dim SpaceCount As Integer = 0

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

    Public Function GetPhraseFromSubSeg(ByVal SubSegID As Integer, ByRef SubSeg As HTML_SUBSEG,
                                        ByRef Phrase() As HTML_PHRASE, ByRef GlobalPhraseCount As Integer) As Boolean
        GetPhraseFromSubSeg = False

        '### Phrase(문장) 나누기 ★ ###
        Dim T_PhraseCountPerSubSeg As Integer = 0
        Dim T_PhrasesCount As Integer = GlobalPhraseCount
        Dim T_SubSeg As HTML_SUBSEG = SubSeg

        For j = 1 To T_SubSeg.Content.Length
            If Mid(T_SubSeg.Content, j, 1) = "." Then

                If j = 1 Then
                    '### 맨 처음인데 .이 찍혀있으면 나누지 않음
                ElseIf Mid(T_SubSeg.Content, j - 1, 1) = ">" Or Mid(T_SubSeg.Content, j - 1, 1) = "." Then
                    '### 맨 처음 글자가 .이거나 바로 앞 글자도 .이었을 경우 나누지 않음
                Else

                    If j = T_SubSeg.Content.Length Then
                        '### 마지막이 .이면 당연히 나눈다!
                        T_PhrasesCount = T_PhrasesCount + 1
                        T_PhraseCountPerSubSeg = T_PhraseCountPerSubSeg + 1
                        Phrase(T_PhrasesCount - 1).Content = Mid(T_SubSeg.Content, 1, j)
                        Phrase(T_PhrasesCount - 1).IndexInSubSeg = T_PhraseCountPerSubSeg - 1
                        Phrase(T_PhrasesCount - 1).ParentSubSegID = SubSegID
                        Phrase(T_PhrasesCount - 1).ParentSegID = T_SubSeg.ParentSegID
                        T_SubSeg.Content = Mid(T_SubSeg.Content, j + 1, T_SubSeg.Content.Length)
                        j = 0

                    ElseIf Mid(T_SubSeg.Content, j + 1, 1) = " " Then
                        '### 다음 글자가 띄어쓰기여도 나눈다!! ★★ (. 다음이 띄어쓰기가 아니면 안 나눈다!!★)
                        T_PhrasesCount = T_PhrasesCount + 1
                        T_PhraseCountPerSubSeg = T_PhraseCountPerSubSeg + 1
                        Phrase(T_PhrasesCount - 1).Content = Mid(T_SubSeg.Content, 1, j + 1)
                        Phrase(T_PhrasesCount - 1).IndexInSubSeg = T_PhraseCountPerSubSeg - 1
                        Phrase(T_PhrasesCount - 1).ParentSubSegID = SubSegID
                        Phrase(T_PhrasesCount - 1).ParentSegID = T_SubSeg.ParentSegID
                        T_SubSeg.Content = Mid(T_SubSeg.Content, j + 2, T_SubSeg.Content.Length)
                        j = 0
                    End If
                End If
            End If
        Next

        T_SubSeg.Content = T_SubSeg.Content.Replace("&nbsp;", "") '### &nbsp; 는 빈 내용이므로 번역이 필요 없으니까!!

        If T_SubSeg.Content = " " Then T_SubSeg.Content = "" '### 띄어쓰기 하나만 남았으면 없애자!

        If T_SubSeg.Content <> "" Then '### 아직 내용이 남아 있으면 새로 추가!!
            T_PhrasesCount = T_PhrasesCount + 1
            T_PhraseCountPerSubSeg = T_PhraseCountPerSubSeg + 1
            Phrase(T_PhrasesCount - 1).Content = SubSeg.Content
            Phrase(T_PhrasesCount - 1).IndexInSubSeg = T_PhraseCountPerSubSeg - 1
            Phrase(T_PhrasesCount - 1).ParentSubSegID = SubSegID
            Phrase(T_PhrasesCount - 1).ParentSegID = SubSeg.ParentSegID
        End If

        SubSeg.PhraseCount = T_PhraseCountPerSubSeg

        GlobalPhraseCount = T_PhrasesCount

        GetPhraseFromSubSeg = True
    End Function

End Module
