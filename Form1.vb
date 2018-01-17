'Provide necessary error handling in the code.
Imports System.Runtime.InteropServices

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnSpellCheck As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents btnGrammarCheck As System.Windows.Forms.Button
    Friend WithEvents lnkReadMe As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnSpellCheck = New System.Windows.Forms.Button()
        Me.btnGrammarCheck = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.lnkReadMe = New System.Windows.Forms.LinkLabel()
        Me.SuspendLayout()
        '
        'btnSpellCheck
        '
        Me.btnSpellCheck.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSpellCheck.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSpellCheck.Location = New System.Drawing.Point(16, 148)
        Me.btnSpellCheck.Name = "btnSpellCheck"
        Me.btnSpellCheck.Size = New System.Drawing.Size(104, 32)
        Me.btnSpellCheck.TabIndex = 1
        Me.btnSpellCheck.Text = "Spelling Check"
        '
        'btnGrammarCheck
        '
        Me.btnGrammarCheck.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnGrammarCheck.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnGrammarCheck.Location = New System.Drawing.Point(152, 148)
        Me.btnGrammarCheck.Name = "btnGrammarCheck"
        Me.btnGrammarCheck.Size = New System.Drawing.Size(104, 32)
        Me.btnGrammarCheck.TabIndex = 2
        Me.btnGrammarCheck.Text = "Grammar Check"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(16, 10)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(352, 120)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = ""
        '
        'lnkReadMe
        '
        Me.lnkReadMe.AutoSize = True
        Me.lnkReadMe.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkReadMe.Location = New System.Drawing.Point(312, 167)
        Me.lnkReadMe.Name = "lnkReadMe"
        Me.lnkReadMe.Size = New System.Drawing.Size(48, 13)
        Me.lnkReadMe.TabIndex = 3
        Me.lnkReadMe.TabStop = True
        Me.lnkReadMe.Text = "ReadMe"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(386, 200)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lnkReadMe, Me.TextBox1, Me.btnGrammarCheck, Me.btnSpellCheck})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximumSize = New System.Drawing.Size(392, 232)
        Me.MinimumSize = New System.Drawing.Size(392, 232)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Subroutine: SpellOrGrammarCheck "

    'Note: Spelling checker only checks spelling, whereas Grammar checker check
    '      Grammar & Spellings.
    Private Sub SpellOrGrammarCheck(ByVal SpellCheckOnly As Boolean)
        Dim objWord As Object       'Create Word document objetc.
        Dim objTempDoc As Object    'Create Temporary document object.
        Dim iData As IDataObject    'To hold data returned from clipboard.

        Try
            'If TextBox is empty, then exit Subroutine.
            If TextBox1.Text = "" Then
                MessageBox.Show("There are no contents to check.", "Spell Checker", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox1.Focus()
                Exit Sub
            End If

            'Instantiate Word application, add temporary document to it and
            'hide instantiated Word application.
            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.visible = False

            'Position Word off the screen, keeps Word invisible throughout.
            objWord.WindowState = 0
            objWord.Top = -3500

            'Copy contents of  TextBox to the clipboard.
            Clipboard.SetDataObject(TextBox1.Text)

            'Perform either spell check or grammar check.
            With objTempDoc
                .Content.Paste()
                .Activate()
                If SpellCheckOnly Then
                    .CheckSpelling()
                Else
                    .CheckGrammar()
                End If

                'Transfer the contents clipboard to the TextBox
                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    TextBox1.Text = CType(iData.GetData(DataFormats.Text), String)
                End If
                .Saved = True
                .Close()
            End With

            'Quit Word.
            objWord.Quit()
            MessageBox.Show("Completed.", "Spell Checker", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'MS Word must be installed.
        Catch COMExcep As COMException
            MessageBox.Show(COMExcep.Message, "Spell Checker", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch Excep As Exception
            MessageBox.Show(Excep.Message, "Spell Checker", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region " Button Click Functionality "

    Private Sub onSpellCheckClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpellCheck.Click
        SpellOrGrammarCheck(True)
    End Sub

    Private Sub onGrammarCheckClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGrammarCheck.Click
        SpellOrGrammarCheck(False)
    End Sub

    Private Sub onReadMeClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkReadMe.LinkClicked
        Dim strReadMeLoc As String

        'Set the path for ReadMe file.
        strReadMeLoc = Application.StartupPath
        strReadMeLoc = Microsoft.VisualBasic.Left(strReadMeLoc, InStr(strReadMeLoc, "\bin")) & "Readme.htm"
        lnkReadMe.LinkVisited = True
        lnkReadMe.VisitedLinkColor = Color.Red
        Process.Start(strReadMeLoc)
    End Sub

#End Region

End Class
