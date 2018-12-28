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
        Me.TBWord = New System.Windows.Forms.TextBox()
        Me.TBST = New System.Windows.Forms.TextBox()
        Me.TBSTSentence = New System.Windows.Forms.TextBox()
        Me.TBTTSentence = New System.Windows.Forms.TextBox()
        Me.TBTT = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.LBCurSentence = New System.Windows.Forms.Label()
        Me.CMS1.SuspendLayout()
        Me.Panel1.SuspendLayout()
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
        Me.TBWord.Location = New System.Drawing.Point(12, 160)
        Me.TBWord.Name = "TBWord"
        Me.TBWord.Size = New System.Drawing.Size(138, 21)
        Me.TBWord.TabIndex = 2
        '
        'TBST
        '
        Me.TBST.BackColor = System.Drawing.Color.Honeydew
        Me.TBST.Location = New System.Drawing.Point(172, 160)
        Me.TBST.Multiline = True
        Me.TBST.Name = "TBST"
        Me.TBST.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBST.Size = New System.Drawing.Size(404, 423)
        Me.TBST.TabIndex = 0
        '
        'TBSTSentence
        '
        Me.TBSTSentence.BackColor = System.Drawing.Color.LightGreen
        Me.TBSTSentence.Location = New System.Drawing.Point(172, 12)
        Me.TBSTSentence.Multiline = True
        Me.TBSTSentence.Name = "TBSTSentence"
        Me.TBSTSentence.ReadOnly = True
        Me.TBSTSentence.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBSTSentence.Size = New System.Drawing.Size(404, 142)
        Me.TBSTSentence.TabIndex = 3
        '
        'TBTTSentence
        '
        Me.TBTTSentence.BackColor = System.Drawing.Color.LightGreen
        Me.TBTTSentence.Location = New System.Drawing.Point(592, 12)
        Me.TBTTSentence.Multiline = True
        Me.TBTTSentence.Name = "TBTTSentence"
        Me.TBTTSentence.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBTTSentence.Size = New System.Drawing.Size(404, 142)
        Me.TBTTSentence.TabIndex = 5
        '
        'TBTT
        '
        Me.TBTT.BackColor = System.Drawing.Color.Honeydew
        Me.TBTT.Location = New System.Drawing.Point(592, 160)
        Me.TBTT.Multiline = True
        Me.TBTT.Name = "TBTT"
        Me.TBTT.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TBTT.Size = New System.Drawing.Size(404, 423)
        Me.TBTT.TabIndex = 6
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.LBCurSentence)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(154, 142)
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
        'MainFrm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1008, 606)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TBTT)
        Me.Controls.Add(Me.TBTTSentence)
        Me.Controls.Add(Me.TBSTSentence)
        Me.Controls.Add(Me.TBST)
        Me.Controls.Add(Me.TBWord)
        Me.Name = "MainFrm"
        Me.Text = "HUFS GSIT CAT"
        Me.CMS1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CMS1 As ContextMenuStrip
    Friend WithEvents 검색ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 닫기ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TBWord As TextBox
    Friend WithEvents TBST As TextBox
    Friend WithEvents TBSTSentence As TextBox
    Friend WithEvents TBTTSentence As TextBox
    Friend WithEvents TBTT As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents LBCurSentence As Label
End Class
