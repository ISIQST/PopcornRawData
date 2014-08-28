Public Class Application
    Implements IDisposable

    Public ReadOnly Property QuasiAddIn() As Boolean
        Get
            Return True
        End Get
    End Property
    Public Property ModuleDescr() As String = "Popcorn Raw Data Extractor"
    Public Property ModuleID() As String = "PopcornRawData"

    Public Sub Initialize2(ByRef qs As Object)
        Try
            'set pointer to QST 
            qst = qs
            'regsiter test class
            qst.QuasiParameters.RegisterTestClassNET(DigDataAnalyzer.SharedTestID, ModuleID, ModuleID & ".DigDataAnalyzer", My.Resources.Icon1, ModuleID & ".ucMain")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If qst IsNot Nothing Then qst.QuasiParameters.UnregisterTestClass(DigDataAnalyzer.SharedTestID)
                qst = Nothing
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
