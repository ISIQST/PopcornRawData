Imports System.Windows.Forms

Public Class ucMain

    Public WithEvents OwnerTest As DigDataAnalyzer

    Private ptrInitSubTests As MethodInvoker = AddressOf OpenTestMenus
    Private Sub lOwnerTest_InitTestForms(ByVal tList As Object) Handles OwnerTest.InitTestForms
        Invoke(ptrInitSubTests, tlist)
    End Sub

    Public Overrides Sub RunTimeInit(ByRef OwnerTestptr As Quasi97.clsQSTTestNET)
        MyBase.RunTimeInit(OwnerTestptr)
        OwnerTest = CType(OwnerTestptr, DigDataAnalyzer)
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ucMain_Disposed(sender As Object, e As System.EventArgs) Handles Me.Disposed
        OwnerTest = Nothing
    End Sub
End Class
