Module modMain
    Public qst As Quasi97.Application

    Public GUITestKeys As New System.Collections.Generic.List(Of String)

    Public Sub OpenTestMenus()
        'ensure the controls for the tests are loaded
        Try
            Dim ct As Object = qst.GetMainModNew.CurTestNet
            With qst.QuasiParameters
                For i As Integer = 0 To GUITestKeys.Count - 1
                    If Not .TestControls.ContainsKey(GUITestKeys(i)) AndAlso .TestListLegacy(GUITestKeys(i)).NETOBJ Then
                        qst.GetMainModNew.CurTestNet = .TestListLegacy(GUITestKeys(i)).TestPtr
                    End If
                Next i
                System.Windows.Forms.Application.DoEvents() 'if logplots, then doevent already occured
                qst.GetMainModNew.CurTestNet = ct
            End With

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Module
