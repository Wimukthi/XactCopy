' -----------------------------------------------------------------------------
' File: src\XactCopy.Storage\Infrastructure\RecoveryStateStore.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.IO
Imports System.Text.Json
Imports XactCopy.Models

Namespace Infrastructure
    ''' <summary>
    ''' Class RecoveryStateStore.
    ''' </summary>
    Public Class RecoveryStateStore
        Private ReadOnly _statePath As String
        Private ReadOnly _serializerOptions As JsonSerializerOptions

        ''' <summary>
        ''' Initializes a new instance.
        ''' </summary>
        Public Sub New(Optional statePath As String = Nothing)
            If String.IsNullOrWhiteSpace(statePath) Then
                Dim root = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XactCopy",
                    "runtime")
                _statePath = Path.Combine(root, "recovery-state.json")
            Else
                _statePath = statePath
            End If

            _serializerOptions = New JsonSerializerOptions() With {
                .WriteIndented = True,
                .PropertyNameCaseInsensitive = True
            }
        End Sub

        ''' <summary>
        ''' Gets or sets StatePath.
        ''' </summary>
        Public ReadOnly Property StatePath As String
            Get
                Return _statePath
            End Get
        End Property

        ''' <summary>
        ''' Computes Load.
        ''' </summary>
        Public Function Load() As RecoveryState
            If Not File.Exists(_statePath) Then
                Return New RecoveryState()
            End If

            Try
                Dim json = File.ReadAllText(_statePath)
                Dim state = JsonSerializer.Deserialize(Of RecoveryState)(json, _serializerOptions)
                Return NormalizeState(state)
            Catch
                Return New RecoveryState()
            End Try
        End Function

        ''' <summary>
        ''' Executes Save.
        ''' </summary>
        Public Sub Save(state As RecoveryState)
            If state Is Nothing Then
                Throw New ArgumentNullException(NameOf(state))
            End If

            Dim normalized = NormalizeState(state)
            Dim directoryPath = Path.GetDirectoryName(_statePath)
            If Not String.IsNullOrWhiteSpace(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            Dim tempPath = _statePath & ".tmp"
            Dim json = JsonSerializer.Serialize(normalized, _serializerOptions)
            File.WriteAllText(tempPath, json)
            File.Move(tempPath, _statePath, overwrite:=True)
        End Sub

        Private Shared Function NormalizeState(state As RecoveryState) As RecoveryState
            If state Is Nothing Then
                Return New RecoveryState()
            End If

            If state.ActiveRun IsNot Nothing AndAlso state.ActiveRun.Options Is Nothing Then
                state.ActiveRun.Options = New CopyJobOptions()
            End If

            Return state
        End Function
    End Class
End Namespace
