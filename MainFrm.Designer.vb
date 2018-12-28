<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainFrm
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainFrm))
        Me.CMS1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.검색ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.닫기ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBWord = New System.Windows.Forms.TextBox()
        Me.TBSTSentence = New System.Windows.Forms.TextBox()
        Me.TBTTSentence = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.LBCurSentence = New System.Windows.Forms.Label()
        Me.RTBST = New System.Windows.Forms.RichTextBox()
        Me.MSMainMenu = New System.Windows.Forms.MenuStrip()
        Me.파일ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.열기ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.종료ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OFD = New System.Windows.Forms.OpenFileDialog()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.TSBFont = New System.Windows.Forms.ToolStripButton()
        Me.FD = New System.Windows.Forms.FontDialog()
        Me.RTBTT = New System.Windows.Forms.RichTextBox()
        Me.CMS1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.MSMainMenu.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CMS1
        '
        Me.CMS1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.검색ToolStripMenuItem, Me.닫기ToolStripMenuItem})
        Me.CMS1.Name = "CMS1"
        Me.CMS1.Size = New System.Drawing.Size(99, 48)
        '
        '검색ToolStripMenuItem
        '
        Me.검색ToolStripMenuItem.Name = "검색ToolStripMenuItem"
        Me.검색ToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.검색ToolStripMenuItem.Text = "검색"
        '
        '닫기ToolStripMenuItem
        '
        Me.닫기ToolStripMenuItem.Name = "닫기ToolStripMenuItem"
        Me.닫기ToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.닫기ToolStripMenuItem.Text = "닫기"
        '
        'TBWord
        '
        Me.TBWord.Location = New System.Drawing.Point(12, 229)
        Me.TBWord.Name = "TBWord"
        Me.TBWord.Size = New System.Drawing.Size(138, 21)
        Me.TBWord.TabIndex = 2
        '
        'TBSTSentence
        '
        Me.TBSTSentence.BackColor = System.Drawing.Color.LightGreen
        Me.TBSTSentence.Location = New System.Drawing.Point(172, 55)
        Me.TBSTSentence.Multiline = True
        Me.TBSTSentence.Name = "TBSTSentence"
        Me.TBSTSentence.ReadOnly = True
        Me.TBSTSentence.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBSTSentence.Size = New System.Drawing.Size(404, 168)
        Me.TBSTSentence.TabIndex = 3
        '
        'TBTTSentence
        '
        Me.TBTTSentence.BackColor = System.Drawing.Color.LightGreen
        Me.TBTTSentence.Location = New System.Drawing.Point(592, 55)
        Me.TBTTSentence.Multiline = True
        Me.TBTTSentence.Name = "TBTTSentence"
        Me.TBTTSentence.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBTTSentence.Size = New System.Drawing.Size(404, 168)
        Me.TBTTSentence.TabIndex = 5
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.LBCurSentence)
        Me.Panel1.Location = New System.Drawing.Point(12, 55)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(154, 99)
        Me.Panel1.TabIndex = 7
        '
        'LBCurSentence
        '
        Me.LBCurSentence.AutoSize = True
        Me.LBCurSentence.Location = New System.Drawing.Point(3, 2)
        Me.LBCurSentence.Name = "LBCurSentence"
        Me.LBCurSentence.Size = New System.Drawing.Size(61, 12)
        Me.LBCurSentence.TabIndex = 0
        Me.LBCurSentence.Text = "현재 문장:"
        '
        'RTBST
        '
        Me.RTBST.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RTBST.Location = New System.Drawing.Point(172, 232)
        Me.RTBST.Name = "RTBST"
        Me.RTBST.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth
        Me.RTBST.Size = New System.Drawing.Size(404, 351)
        Me.RTBST.TabIndex = 8
        Me.RTBST.Text = ""
        '
        'MSMainMenu
        '
        Me.MSMainMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.파일ToolStripMenuItem})
        Me.MSMainMenu.Location = New System.Drawing.Point(0, 0)
        Me.MSMainMenu.Name = "MSMainMenu"
        Me.MSMainMenu.Size = New System.Drawing.Size(1008, 24)
        Me.MSMainMenu.TabIndex = 9
        Me.MSMainMenu.Text = "MenuStrip1"
        '
        '파일ToolStripMenuItem
        '
        Me.파일ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.열기ToolStripMenuItem, Me.종료ToolStripMenuItem})
        Me.파일ToolStripMenuItem.Name = "파일ToolStripMenuItem"
        Me.파일ToolStripMenuItem.Size = New System.Drawing.Size(43, 20)
        Me.파일ToolStripMenuItem.Text = "파일"
        '
        '열기ToolStripMenuItem
        '
        Me.열기ToolStripMenuItem.Name = "열기ToolStripMenuItem"
        Me.열기ToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.열기ToolStripMenuItem.Text = "열기"
        '
        '종료ToolStripMenuItem
        '
        Me.종료ToolStripMenuItem.Name = "종료ToolStripMenuItem"
        Me.종료ToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.종료ToolStripMenuItem.Text = "종료"
        '
        'OFD
        '
        Me.OFD.FileName = "OpenFileDialog1"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSBFont})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 24)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1008, 25)
        Me.ToolStrip1.TabIndex = 10
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'TSBFont
        '
        Me.TSBFont.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSBFont.Image = CType(resources.GetObject("TSBFont.Image"), System.Drawing.Image)
        Me.TSBFont.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSBFont.Name = "TSBFont"
        Me.TSBFont.Size = New System.Drawing.Size(23, 22)
        Me.TSBFont.Text = "ToolStripButton1"
        '
        'RTBTT
        '
        Me.RTBTT.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RTBTT.Location = New System.Drawing.Point(592, 229)
        Me.RTBTT.Name = "RTBTT"
        Me.RTBTT.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth
        Me.RTBTT.Size = New System.Drawing.Size(404, 351)
        Me.RTBTT.TabIndex = 11
        Me.RTBTT.Text = ""
        '
        'MainFrm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1008, 606)
        Me.Controls.Add(Me.RTBTT)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.RTBST)
        Me.Controls.Add(Me.MSMainMenu)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TBTTSentence)
        Me.Controls.Add(Me.TBSTSentence)
        Me.Controls.Add(Me.TBWord)
        Me.MainMenuStrip = Me.MSMainMenu
        Me.Name = "MainFrm"
        Me.Text = "HUFS GSIT CAT"
        Me.CMS1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.MSMainMenu.ResumeLayout(False)
        Me.MSMainMenu.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CMS1 As ContextMenuStrip
    Friend WithEvents 검색ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 닫기ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TBWord As TextBox
    Friend WithEvents TBSTSentence As TextBox
    Friend WithEvents TBTTSentence As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents LBCurSentence As Label
    Friend WithEvents RTBST As RichTextBox
    Friend WithEvents MSMainMenu As MenuStrip
    Friend WithEvents 파일ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 열기ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 종료ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OFD As OpenFileDialog
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents TSBFont As ToolStripButton
    Friend WithEvents FD As FontDialog
    Friend WithEvents RTBTT As RichTextBox
End Class
