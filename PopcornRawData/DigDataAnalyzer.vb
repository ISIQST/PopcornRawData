Option Strict On
Public Class DigDataAnalyzer
    Inherits Quasi97.clsQSTTestNET

    Public Shared SharedTestID$ = "DigData Analyzer"
    Public Property PopTestUserName As String = "1"

    Private WithEvents popTest As NoiseApp.clsPopcornTest
    Private rslt As Quasi97.ResultNet
    Private Class NV
        Public RMS$ = "rms (uV)"
    End Class
    Private myNV As New NV

    Public Event InitTestForms(ByVal tlist() As String)

    Public Overrides Function CheckRecords(NewDBase As String) As System.Collections.Generic.List(Of Short)
        Return New List(Of Short)({1})
        'no database support
    End Function

    Public Overrides Sub ClearResults(Optional doRefreshPlot As Boolean = False)
        colResults.Clear()
        MyBase.RaiseResultsCleared(doRefreshPlot)
    End Sub

    Public Overrides ReadOnly Property ContainsGraph As Short
        Get
            Return -1
        End Get
    End Property

    Public Overrides ReadOnly Property ContainsResultPerCycle As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property DualChannelCapable As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property FeatureVector As UInteger
        Get
            Return 0
        End Get
    End Property

    Public Overrides Sub RemoveRecord()
        'no database support
    End Sub

    Public Overrides Sub RestoreParameters()
        'no database support
    End Sub

    Public Overrides Sub RunTest()
        Dim tShell As Quasi97.TestShell = Nothing
        Dim FoundMatch As Boolean = False
        Try
            'clear results
            colResults.Clear()

            'find the popcorn test that user wants to run
            For Each tShell In qst.QuasiParameters.TestListLegacy.Values
                If (tShell.TestID = NoiseApp.clsPopcornTest.SharedTestID) And (tShell.UserName = PopTestUserName) Then
                    FoundMatch = True
                    Exit For
                End If
            Next

            If Not foundmatch Then Return 'test was not found
            RaiseEvent InitTestForms({tShell.Key})

            'connect event handler
            popTest = CType(tShell.TestPtr, NoiseApp.clsPopcornTest)

            'create new result for this test
            rslt = New Quasi97.ResultNet
            colResults.Add(rslt)
            'document parameters
            rslt.AddParameters(colParameters)

            'run popcorn test, eventhandler will capture the results
            Quasi97.TestShell.RunTestNET(CType(popTest, Quasi97.clsQSTTestNET))

            'analyze data and generate some results
            AnalyzeRawData(rslt)
            MyBase.RaiseNewDataAvailable(colResults.Count, 1, -1, 0)
            MyBase.RaiseNewResultsAvailable({0}) 'notify the test menu that new results are available

            'add information about the test
            qst.GradingParameters.GradeTestNet2(Me, 0)
            rslt.AddInfo2(Me, 0, qst.QuasiParameters.CurInfo)
            MyBase.RaiseNewInfoAvailable()

            If Not AbortTest Then
                If Not qst.QuasiParameters.ProductionMode Then qst.QuasiParameters.LogResults2(Me, 0)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            popTest = Nothing 'disconnect from event handler

        End Try

    End Sub

    Private Sub AnalyzeRawData(myrslt As Quasi97.ResultNet)
        Try
            'for example calculate rms
            With myrslt.Data
                Cycles = .Columns.Count - 1
                For i = 1 To .Columns.Count - 1
                    If i > 1 Then myrslt.Results.Columns.Add() 'column for the first cycle is added when rslt is created
                    Dim sum# = 0
                    For j = 0 To .Rows.Count - 1
                        sum += CDbl(.Rows(j).Item(i)) ^ 2
                    Next
                    myrslt.AddResult2(myNV.RMS, Math.Sqrt(sum / .Rows.Count), CShort(i), True)
                Next
            End With

            myrslt.CalcStats2(False)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Overrides Sub SetDBase(ByRef NewDBase As String, Optional ByRef voidParam As Object = Nothing)
        'no database support
    End Sub

    Public Overrides Sub StoreParameters()
        'no database support
    End Sub

    Public Overrides ReadOnly Property TestID As String
        Get
            Return SharedTestID
        End Get
    End Property

    Public Sub New()
        'register parameters
        colParameters.Add(New Quasi97.clsTestParam("PopTestUserName", "PopTestUserName", Me, GetType(String), False, False))

        'register results
        RegisterResults2(CType(myNV, Object))
    End Sub

    'this is the function that captures raw data from popcorn test
    Private Sub popTest_NewRawDataAvailable(Events(,) As Integer, Points As Integer, xdat() As Single, ydat() As Single, MaxX As Single, MaxY As Single) Handles popTest.NewRawDataAvailable
        Try
            If Events Is Nothing Then Return 'dummy run to clear the data

            Dim Cycles% = 1 + UBound(Events, 2)
            Dim Pts% = Points

            'add columns
            With rslt.Data.Columns
                .Clear()
                .Add("Time (us)")

                For i = 1 To Cycles
                    .Add("C" & i.ToString & " Amp (uV)")
                Next
            End With

            'add rows
            For i = 0 To Points - 1
                Dim dr As DataRow = rslt.Data.Rows.Add()
                dr(0) = xdat(i)
                Dim jj As Integer = i
                For j = 1 To Cycles
                    dr(j) = ydat(jj)
                    jj += Points
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
