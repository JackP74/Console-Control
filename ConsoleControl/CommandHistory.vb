Imports System
Imports System.Collections
Imports System.Text

Friend Class CommandHistory
    Private currentPosn As Integer
    Private CLastCommand As String
    Private commandHistory As ArrayList = New ArrayList()

    Friend Sub New()
    End Sub

    Friend Sub Add(ByVal command As String)
        If command <> lastCommand Then
            commandHistory.Add(command)
            CLastCommand = command
            currentPosn = commandHistory.Count
        End If
    End Sub

    Friend Function DoesPreviousCommandExist() As Boolean
        Return currentPosn > 0
    End Function

    Friend Function DoesNextCommandExist() As Boolean
        Return currentPosn < commandHistory.Count - 1
    End Function

    Friend Function GetPreviousCommand() As String
        CLastCommand = CStr(commandHistory(System.Threading.Interlocked.Decrement(currentPosn)))
        Return lastCommand
    End Function

    Friend Function GetNextCommand() As String
        CLastCommand = CStr(commandHistory(System.Threading.Interlocked.Increment(currentPosn)))
        Return lastCommand
    End Function

    Friend ReadOnly Property LastCommand As String
        Get
            Return CLastCommand
        End Get
    End Property

    Friend Function GetCommandHistory() As String()
        Return CType(commandHistory.ToArray(GetType(String)), String())
    End Function

    Friend Function ClearHistory() As String
        commandHistory = New ArrayList()
        Return String.Empty
    End Function
End Class
