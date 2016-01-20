Imports System.Data.SqlClient
Imports System.Data


Public Class Catalogs
    Dim cnn As New SqlConnection(ConfigurationManager.ConnectionStrings("MyConnectionString").ConnectionString)

    Private Sub Catalogs_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Dim m As New MenuP()
        m.Show()
    End Sub

    Private Sub UnitType_btn_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles UnitType_btn.Click
        Cat_insert(1, UnitType_txt.Text)
        UnitType_txt.Text = ""
    End Sub

    Private Sub Cat_insert(ByVal Tipo As Integer, ByVal Value As String)
        conectado()
        Using cmd As New SqlCommand("Catalog_Insert", cnn)
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter("@Tipo", SqlDbType.Int)
            param1.Value = Tipo
            Dim param2 As New SqlParameter("@Value", SqlDbType.NVarChar)
            param2.Value = Value

            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)

            If cmd.ExecuteNonQuery() Then
                MsgBox("Successful", MsgBoxStyle.Information, "Alta")
                carga_Combos(Tipo)
            Else
                MsgBox("Error", MsgBoxStyle.Information, "Alta")
            End If
        End Using
        desconectado()
    End Sub
    Private Sub Cat_insert2(ByVal Tipo As Integer, ByVal Value As String, ByVal value2 As Integer)
        conectado()
        Using cmd As New SqlCommand("Catalog_Insert", cnn)
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter("@Tipo", SqlDbType.Int)
            param1.Value = Tipo
            Dim param2 As New SqlParameter("@Value", SqlDbType.NVarChar)
            param2.Value = Value
            Dim param3 As New SqlParameter("@Value2", SqlDbType.Int)
            param3.Value = value2

            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)

            If cmd.ExecuteNonQuery() Then
                MsgBox("Successful", MsgBoxStyle.Information, "Alta")
            Else
                MsgBox("Error", MsgBoxStyle.Information, "Alta")
            End If
        End Using
        desconectado()
    End Sub

    Private Sub Catalogs_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded

    End Sub

    Private Sub carga_Combos(ByVal tipo As Integer)

        Unittype_cb.Items.Clear()

        Dim dt As New DataTable
        Dim ds As New DataSet
        conectado()
        Using cmd As New SqlCommand("Catalog_Read", cnn)
            cmd.CommandType = CommandType.StoredProcedure

            Dim param1 As New SqlParameter("@Tipo", SqlDbType.Int)
            param1.Value = tipo
            cmd.Parameters.Add(param1)
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            While dr.Read()
                Unittype_cb.Items.Add(dr.Item(0))
            End While
            'dt.Load(dr)
            'ds.Tables.Add(dt)
        End Using
        desconectado()
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
    Private Sub Unit_GotFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles Unit.GotFocus
        carga_Combos(1)
    End Sub

    Private Sub Unit_btn_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Unit_btn.Click
        Cat_insert2(2, Unittype_cb.SelectedItem, Unit_txt.Text)
        Unit_txt.Text = ""
        Unittype_cb.SelectedIndex = -1
    End Sub
End Class
