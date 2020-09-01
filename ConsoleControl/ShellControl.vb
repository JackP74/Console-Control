Imports System.Drawing
Imports System.Windows.Forms

Public Class ShellControl
    Inherits UserControl

    Private shellTextBox As ShellTextBox
    Public Event CommandEntered As EventCommandEntered
    Private components As System.ComponentModel.Container = Nothing

    Public Sub New()
        InitializeComponent()
    End Sub

    Friend Sub FireCommandEntered(ByVal command As String)
        OnCommandEntered(command)
    End Sub

    Protected Overridable Sub OnCommandEntered(ByVal command As String)
        RaiseEvent CommandEntered(command, New CommandEnteredEventArgs(command))
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If components IsNot Nothing Then
                components.Dispose()
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub

    Public Property ShellTextForeColor As Color
        Get
            Return If(shellTextBox IsNot Nothing, shellTextBox.ForeColor, Color.Green)
        End Get
        Set(ByVal value As Color)
            If shellTextBox IsNot Nothing Then shellTextBox.ForeColor = value
        End Set
    End Property

    Public Property ShellTextBackColor As Color
        Get
            Return If(shellTextBox IsNot Nothing, shellTextBox.BackColor, Color.Black)
        End Get
        Set(ByVal value As Color)
            If shellTextBox IsNot Nothing Then shellTextBox.BackColor = value
        End Set
    End Property

    Public Property ShellTextFont As Font
        Get
            Return If(shellTextBox IsNot Nothing, shellTextBox.Font, New Font("Tahoma", 8))
        End Get
        Set(ByVal value As Font)
            If shellTextBox IsNot Nothing Then shellTextBox.Font = value
        End Set
    End Property

    Public Sub Clear()
        shellTextBox.Clear()
    End Sub

    Public Sub WriteText(ByVal text As String)
        shellTextBox.WriteText(text)
    End Sub

    Public Function GetCommandHistory() As String()
        Return shellTextBox.GetCommandHistory()
    End Function

    Public Property Prompt As String
        Get
            Return shellTextBox.Prompt
        End Get
        Set(ByVal value As String)
            shellTextBox.Prompt = value
        End Set
    End Property

    Private Sub InitializeComponent()
        Me.shellTextBox = New ShellTextBox()
        Me.SuspendLayout()
        Me.shellTextBox.AcceptsReturn = True
        Me.shellTextBox.AcceptsTab = True
        Me.shellTextBox.BackColor = System.Drawing.Color.Black
        Me.shellTextBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.shellTextBox.ForeColor = System.Drawing.Color.LawnGreen
        Me.shellTextBox.Location = New System.Drawing.Point(0, 0)
        Me.shellTextBox.Multiline = True
        Me.shellTextBox.Margin = New Padding(0)
        Me.shellTextBox.Name = "shellTextBox"
        Me.shellTextBox.Prompt = ">>>"
        Me.shellTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.shellTextBox.BackColor = System.Drawing.Color.Black
        Me.shellTextBox.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte((0))))
        Me.shellTextBox.ForeColor = System.Drawing.Color.LawnGreen
        Me.shellTextBox.Size = New System.Drawing.Size(232, 216)
        Me.shellTextBox.TabIndex = 0
        Me.shellTextBox.Text = ""
        Me.Controls.Add(Me.shellTextBox)
        Me.Name = "ShellControl"
        Me.Size = New System.Drawing.Size(232, 216)
        Me.Padding = New Padding(0)
        Me.Margin = New Padding(0)
        Me.ResumeLayout(False)
    End Sub
End Class

Public Class CommandEnteredEventArgs
    Inherits EventArgs

    Private Ccommand As String

    Public Sub New(ByVal command As String)
        Me.Ccommand = command
    End Sub

    Public ReadOnly Property Command As String
        Get
            Return Ccommand
        End Get
    End Property
End Class

Public Delegate Sub EventCommandEntered(ByVal sender As Object, ByVal e As CommandEnteredEventArgs)