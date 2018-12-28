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
        Me.CMS1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.검색ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.닫기ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBSTSentence = New System.Windows.Forms.TextBox()
        Me.TBTTSentence = New System.Windows.Forms.TextBox()
        Me.MSMainMenu = New System.Windows.Forms.MenuStrip()
        Me.파일ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.열기ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.저장ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.종료ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OFD = New System.Windows.Forms.OpenFileDialog()
        Me.FD = New System.Windows.Forms.FontDialog()
        Me.WBTT = New System.Windows.Forms.WebBrowser()
        Me.LVST = New System.Windows.Forms.ListView()
        Me.CMS1.SuspendLayout()
        Me.MSMainMenu.SuspendLayout()
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
        'TBSTSentence
        '
        Me.TBSTSentence.BackColor = System.Drawing.Color.LightGreen
        Me.TBSTSentence.Location = New System.Drawing.Point(12, 27)
        Me.TBSTSentence.Multiline = True
        Me.TBSTSentence.Name = "TBSTSentence"
        Me.TBSTSentence.ReadOnly = True
        Me.TBSTSentence.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBSTSentence.Size = New System.Drawing.Size(486, 151)
        Me.TBSTSentence.TabIndex = 3
        '
        'TBTTSentence
        '
        Me.TBTTSentence.AcceptsTab = True
        Me.TBTTSentence.BackColor = System.Drawing.Color.LightGreen
        Me.TBTTSentence.Location = New System.Drawing.Point(510, 27)
        Me.TBTTSentence.Multiline = True
        Me.TBTTSentence.Name = "TBTTSentence"
        Me.TBTTSentence.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBTTSentence.Size = New System.Drawing.Size(486, 151)
        Me.TBTTSentence.TabIndex = 5
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
        Me.파일ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.열기ToolStripMenuItem, Me.저장ToolStripMenuItem, Me.종료ToolStripMenuItem})
        Me.파일ToolStripMenuItem.Name = "파일ToolStripMenuItem"
        Me.파일ToolStripMenuItem.Size = New System.Drawing.Size(43, 20)
        Me.파일ToolStripMenuItem.Text = "파일"
        '
        '열기ToolStripMenuItem
        '
        Me.열기ToolStripMenuItem.Name = "열기ToolStripMenuItem"
        Me.열기ToolStripMenuItem.ShortcutKeyDisplayString = ""
        Me.열기ToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.열기ToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.열기ToolStripMenuItem.Text = "열기"
        '
        '저장ToolStripMenuItem
        '
        Me.저장ToolStripMenuItem.Name = "저장ToolStripMenuItem"
        Me.저장ToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.저장ToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.저장ToolStripMenuItem.Text = "저장"
        '
        '종료ToolStripMenuItem
        '
        Me.종료ToolStripMenuItem.Name = "종료ToolStripMenuItem"
        Me.종료ToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.종료ToolStripMenuItem.Text = "종료"
        '
        'OFD
        '
        Me.OFD.FileName = "OpenFileDialog1"
        '
        'WBTT
        '
        Me.WBTT.Location = New System.Drawing.Point(670, 184)
        Me.WBTT.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WBTT.Name = "WBTT"
        Me.WBTT.Size = New System.Drawing.Size(326, 351)
        Me.WBTT.TabIndex = 15
        '
        'LVST
        '
        Me.LVST.Location = New System.Drawing.Point(12, 184)
        Me.LVST.Name = "LVST"
        Me.LVST.Size = New System.Drawing.Size(646, 351)
        Me.LVST.TabIndex = 17
        Me.LVST.UseCompatibleStateImageBehavior = False
        '
        'MainFrm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1008, 544)
        Me.Controls.Add(Me.LVST)
        Me.Controls.Add(Me.WBTT)
        Me.Controls.Add(Me.MSMainMenu)
        Me.Controls.Add(Me.TBTTSentence)
        Me.Controls.Add(Me.TBSTSentence)
        Me.MainMenuStrip = Me.MSMainMenu
        Me.Name = "MainFrm"
        Me.Text = "HUFS GSIT CAT"
        Me.CMS1.ResumeLayout(False)
        Me.MSMainMenu.ResumeLayout(False)
        Me.MSMainMenu.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CMS1 As ContextMenuStrip
    Friend WithEvents 검색ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 닫기ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TBSTSentence As TextBox
    Friend WithEvents TBTTSentence As TextBox
    Friend WithEvents MSMainMenu As MenuStrip
    Friend WithEvents 파일ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 열기ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 종료ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OFD As OpenFileDialog
    Friend WithEvents FD As FontDialog
    Friend WithEvents 저장ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents WBTT As WebBrowser
    Friend WithEvents LVST As ListView
End Class
