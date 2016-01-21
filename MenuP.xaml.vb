Imports System.Data.SqlClient
Imports System.Data

Public Class MenuP
    Dim cnn As New SqlConnection(ConfigurationManager.ConnectionStrings("MyConnectionString").ConnectionString)

    Private Sub Transformer_btn_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Transformer_btn.Click
        Dim trans As New Transformer()
        trans.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Button4.Click
        Application.Current.Shutdown()
        desconectado()
    End Sub

    Private Sub OilQuality_btn_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles OilQuality_btn.Click
        Dim oil As New OilQuality()
        oil.Show()
        Me.Hide()
    End Sub

    Private Sub Catalogs_btn_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Catalogs_btn.Click
        Dim cat As New Catalogs()
        cat.Show()
        Me.Hide()
    End Sub

    Private Sub MenuP_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Application.Current.Shutdown()
        desconectado()
    End Sub

    Private Sub MenuP_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If conectado() = False Then
            Application.Current.Shutdown()
        End If

    End Sub

    Public Function desconectado()
        Try
            If cnn.State = ConnectionState.Open Then
                cnn.Close()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function conectado()
        Try
            If cnn.State = ConnectionState.Closed Then
                cnn.Open()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            cnn.Close()
            Return False
        End Try
    End Function

    Private Sub gauge_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles gauge.Click
        Dim gg As New gauge()
        gg.Show()
        Me.Hide()
    End Sub
End Class
