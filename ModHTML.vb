Module ModHTML
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
End Module
