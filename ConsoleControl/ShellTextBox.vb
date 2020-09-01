Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Windows.Forms
Imports System.Text
Imports System.Diagnostics

Friend Class ShellTextBox
    Inherits TextBox

    Private CPrompt As String = ">>>"
    Private commandHistory As CommandHistory = New CommandHistory()
    Private components As System.ComponentModel.Container = Nothing

    Friend Sub New()
        InitializeComponent()
        printPrompt()
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If components IsNot Nothing Then
                components.Dispose()
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Select Case m.Msg
            Case &H302, &H300, &HC
                If Not IsCaretAtWritablePosition() Then MoveCaretToEndOfText()
            Case &H303
                Return
        End Select

        MyBase.WndProc(m)
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.BackColor = System.Drawing.Color.Black
        Me.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ForeColor = System.Drawing.Color.LawnGreen
        Me.Location = New System.Drawing.Point(0, 0)
        Me.BorderStyle = Windows.Forms.BorderStyle.None
        Me.MaxLength = 0
        Me.Multiline = True
        Me.Name = "shellTextBox"
        Me.AcceptsTab = True
        Me.AcceptsReturn = True
        Me.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.Size = New System.Drawing.Size(400, 176)
        Me.TabIndex = 0
        Me.Text = ""
        AddHandler Me.KeyPress, AddressOf Me.shellTextBox_KeyPress
        AddHandler Me.KeyDown, AddressOf ShellControl_KeyDown
        Me.Name = "ShellTextBox"
        Me.Size = New System.Drawing.Size(400, 176)
        Me.ResumeLayout(False)
    End Sub

    Private Sub printPrompt()
        Dim currentText As String = Me.Text
        If currentText.Length <> 0 AndAlso currentText(currentText.Length - 1) <> vbLf Then printLine()
        Me.AddText(prompt)
    End Sub

    Private Sub printLine()
        Me.AddText(System.Environment.NewLine)
    End Sub

    Private Sub shellTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = ChrW(8) AndAlso IsCaretJustBeforePrompt() Then
            e.Handled = True
            Return
        End If

        If IsTerminatorKey(e.KeyChar) Then
            e.Handled = True
            Dim currentCommand As String = GetTextAtPrompt()

            If currentCommand.Length <> 0 Then
                printLine()
                CType(Me.Parent, ShellControl).FireCommandEntered(currentCommand)
                commandHistory.Add(currentCommand)
            End If

            printPrompt()
        End If
    End Sub

    Private Sub ShellControl_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If Not IsCaretAtWritablePosition() AndAlso Not (e.Control OrElse IsTerminatorKey(e.KeyCode)) Then
            MoveCaretToEndOfText()
        End If

        If e.KeyCode = Keys.Left AndAlso IsCaretJustBeforePrompt() Then
            e.Handled = True
        ElseIf e.KeyCode = Keys.Down Then

            If commandHistory.DoesNextCommandExist() Then
                ReplaceTextAtPrompt(commandHistory.GetNextCommand())
            End If

            e.Handled = True
        ElseIf e.KeyCode = Keys.Up Then

            If commandHistory.DoesPreviousCommandExist() Then
                ReplaceTextAtPrompt(commandHistory.GetPreviousCommand())
            End If

            e.Handled = True
        ElseIf e.KeyCode = Keys.Right Then
            Dim currentTextAtPrompt As String = GetTextAtPrompt()
            Dim lastCommand As String = commandHistory.LastCommand

            If lastCommand IsNot Nothing AndAlso (currentTextAtPrompt.Length = 0 OrElse lastCommand.StartsWith(currentTextAtPrompt)) Then

                If lastCommand.Length > currentTextAtPrompt.Length Then
                    Me.AddText(lastCommand(currentTextAtPrompt.Length).ToString())
                End If
            End If
        End If
    End Sub

    Private Function GetCurrentLine() As String
        If Me.Lines.Length > 0 Then
            Return CStr(Me.Lines.GetValue(Me.Lines.GetLength(0) - 1))
        Else
            Return ""
        End If
    End Function

    Private Function GetTextAtPrompt() As String
        Return GetCurrentLine().Substring(prompt.Length)
    End Function

    Private Sub ReplaceTextAtPrompt(ByVal text As String)
        Dim currentLine As String = GetCurrentLine()
        Dim charactersAfterPrompt As Integer = currentLine.Length - prompt.Length

        If charactersAfterPrompt = 0 Then
            Me.AddText(text)
        Else
            Me.[Select](Me.TextLength - charactersAfterPrompt, charactersAfterPrompt)
            Me.SelectedText = text
        End If
    End Sub

    Private Function IsCaretAtCurrentLine() As Boolean
        Return Me.TextLength - Me.SelectionStart <= GetCurrentLine().Length
    End Function

    Private Sub MoveCaretToEndOfText()
        Me.SelectionStart = Me.TextLength
        Me.ScrollToCaret()
    End Sub

    Private Function IsCaretJustBeforePrompt() As Boolean
        Return IsCaretAtCurrentLine() AndAlso GetCurrentCaretColumnPosition() = prompt.Length
    End Function

    Private Function GetCurrentCaretColumnPosition() As Integer
        Dim currentLine As String = GetCurrentLine()
        Dim currentCaretPosition As Integer = Me.SelectionStart
        Return (currentCaretPosition - Me.TextLength + currentLine.Length)
    End Function

    Private Function IsCaretAtWritablePosition() As Boolean
        Return IsCaretAtCurrentLine() AndAlso GetCurrentCaretColumnPosition() >= prompt.Length
    End Function

    Private Sub SetPromptText(ByVal val As String)
        Dim currentLine As String = GetCurrentLine()
        CPrompt = val
    End Sub

    Public Property Prompt As String
        Get
            Return CPrompt
        End Get
        Set(ByVal value As String)
            SetPromptText(value)
        End Set
    End Property

    Public Function GetCommandHistory() As String()
        Return commandHistory.GetCommandHistory()
    End Function

    Public Sub WriteText(ByVal text As String)
        Me.AddText(text)
    End Sub

    Private Function IsTerminatorKey(ByVal key As Keys) As Boolean
        Return key = Keys.Enter
    End Function

    Private Function IsTerminatorKey(ByVal keyChar As Char) As Boolean
        Return (CInt(AscW(keyChar))) = 13
    End Function

    Private Sub AddText(ByVal text As String)
        Me.Text += text
        MoveCaretToEndOfText()
    End Sub
End Class
