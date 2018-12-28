Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class MainFrm
    Dim g_FontMain As New Font("맑은 고딕", 12, FontStyle.Regular)
    Dim g_FontSentence As New Font("맑은 고딕", 11, FontStyle.Regular)
    Dim g_SentenceID(1000) As Long
    Dim g_NumSentences As Long
    Dim g_sTransSentence(1000) As String    '총 1000 문장
    Dim g_nCurSentID As Long

    Dim g_PrevSTCursor As Long

    Public Sub UpdateTT()
        RTBTT.Text = Nothing

        If g_NumSentences = 0 Then Exit Sub

        For i = 0 To g_NumSentences - 1
            RTBTT.Text = RTBTT.Text & g_sTransSentence(i) & vbCrLf
        Next
    End Sub

    Public Sub DevideSentences(ByVal sText As String)

        If sText.Length = 0 Then Exit Sub

        g_NumSentences = 0
        For i = 1 To sText.Length
            If Mid(sText, i, 1) = "." Or Mid(sText, i, 2) = vbCrLf Then
                g_NumSentences = g_NumSentences + 1
            End If
        Next

        Dim nTempNumSentences As Long = 0
        For i = 1 To sText.Length
            If Mid(sText, i, 1) = "." Then
                nTempNumSentences = nTempNumSentences + 1
                g_SentenceID(nTempNumSentences - 1) = i
            ElseIf Mid(sText, i, 2) = vbCrLf Then
                nTempNumSentences = nTempNumSentences + 1
                g_SentenceID(nTempNumSentences - 1) = i + 1
            End If
        Next

    End Sub

    Public Function GetSentenceID(ByVal CurSelStart As Long) As Long
        GetSentenceID = 0

        For i = 0 To g_NumSentences
            If CurSelStart <= g_SentenceID(i) Then
                GetSentenceID = i
                Exit For
            End If
        Next
    End Function

    Public Function GetSentenceText(ByVal sText As String, ByVal SentenceID As Long) As String
        GetSentenceText = Nothing

        If SentenceID = 0 Then
            GetSentenceText = Trim(Mid(sText, 1, g_SentenceID(SentenceID)))
        Else
            GetSentenceText = Trim(Mid(sText, g_SentenceID(SentenceID - 1) + 1, g_SentenceID(SentenceID) - g_SentenceID(SentenceID - 1)))
        End If

    End Function

    Public Function FindWordInText(ByVal sText As String, ByVal SelStart As Long) As String
        FindWordInText = Nothing

        If sText.Length = 0 Then Exit Function

        Dim nWordSelStart As Long
        Dim nWordSelEnd As Long
        Dim nWordSelLen As Long

        If SelStart = 0 Then
            nWordSelStart = 1
        Else
            nWordSelStart = 1
            For i = SelStart To 1 Step -1
                If Mid(sText, i, 1) = " " Or Mid(sText, i, 1) = vbTab Or Mid(sText, i, 1) = vbLf Then
                    nWordSelStart = i + 1
                    Exit For
                End If
            Next
        End If

        nWordSelEnd = sText.Length + 1
        For i = SelStart + 1 To sText.Length
            If Mid(sText, i, 1) = " " Or Mid(sText, i, 1) = vbTab Or Mid(sText, i, 1) = vbLf Then
                nWordSelEnd = i
                Exit For
            End If
        Next

        nWordSelLen = nWordSelEnd - nWordSelStart
        If nWordSelLen < 0 Then nWordSelLen = 0

        FindWordInText = Mid(sText, nWordSelStart, nWordSelLen)

    End Function

    Public Sub FindWordAndSentences()
        g_PrevSTCursor = RTBST.SelectionStart

        TBWord.Text = FindWordInText(RTBST.Text, g_PrevSTCursor)
        g_nCurSentID = GetSentenceID(g_PrevSTCursor)
        TBSTSentence.Text = GetSentenceText(RTBST.Text, g_nCurSentID)

        LBCurSentence.Text = "현재 문장: " & g_nCurSentID + 1 & " / " & g_NumSentences

        TBTTSentence.Text = g_sTransSentence(g_nCurSentID)
    End Sub

    Private Sub MainFrm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        End
    End Sub

    Private Sub MainFrm_Load(sender As Object, e As EventArgs) Handles Me.Load
        RTBST.Font = g_FontMain
        RTBST.ContextMenuStrip = CMS1
        RTBTT.Font = g_FontMain

        TBSTSentence.Font = g_FontSentence
        TBTTSentence.Font = g_FontSentence
    End Sub

    Private Sub RTBST_KeyUp(sender As Object, e As KeyEventArgs) Handles RTBST.KeyUp
        If RTBST.Text.Length = 0 Then Exit Sub

        FindWordAndSentences()

        If e.KeyCode = Keys.F1 Then
            TBTTSentence.Focus()
        End If
    End Sub
    Private Sub RTBST_MouseUp(sender As Object, e As MouseEventArgs) Handles RTBST.MouseUp
        If RTBST.Text.Length = 0 Then Exit Sub

        FindWordAndSentences()
    End Sub

    Private Sub RTBST_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RTBST.KeyPress
        If e.KeyChar = Convert.ToChar(1) Then
            RTBST.SelectAll()
        End If
    End Sub

    Private Sub RTBST_TextChanged(sender As Object, e As EventArgs) Handles RTBST.TextChanged
        DevideSentences(RTBST.Text)
    End Sub

    Private Sub TBTTSentence_KeyDown(sender As Object, e As KeyEventArgs) Handles TBTTSentence.KeyDown
        If e.KeyCode = Keys.Enter Then
            g_sTransSentence(g_nCurSentID) = TBTTSentence.Text
            e.SuppressKeyPress = True

            UpdateTT()

            RTBST.Focus()
            RTBST.SelectionStart = g_PrevSTCursor
        ElseIf e.KeyCode = Keys.Escape Then
            RTBST.Focus()
            RTBST.SelectionStart = g_PrevSTCursor
        End If
    End Sub

    Private Sub 열기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 열기ToolStripMenuItem.Click
        OFD.Title = "파일 열기"
        OFD.Filter = "워드 파일(*.docx)|*.docx"
        OFD.FileName = Application.ExecutablePath
        OFD.ShowDialog()

        If OFD.FileName IsNot Nothing Then
            Dim objWord As New Word.Application
            Dim objDoc As Word.Document

            Try
                objDoc = objWord.Documents.Open(OFD.FileName)
                objDoc.Content.Copy()
                RTBST.Paste()
                RTBTT.Paste()
                RTBST.ZoomFactor = 0.8
                RTBTT.ZoomFactor = 0.8
                objDoc.Close()
                objWord.Quit()
                objWord = Nothing
            Catch ex As Exception
                MsgBox("오류")
            End Try
        End If
    End Sub

    Private Sub TSBFont_Click(sender As Object, e As EventArgs) Handles TSBFont.Click

        FD.ShowDialog()

        RTBST.SelectionFont = FD.Font
    End Sub
End Class

