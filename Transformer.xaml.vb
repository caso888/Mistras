Imports System.Data.SqlClient
Imports System.Data
Class Transformer

    Dim cnn As New SqlConnection(ConfigurationManager.ConnectionStrings("MyConnectionString").ConnectionString)

    Private Sub Transformer_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Dim m As New MenuP()
        m.Show()
    End Sub

    Private Sub Transformer_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        'ComboBoxEdit8.SelectedIndex = 0
        'ComboBoxEdit8.Items.Add("[Unit]")
        carga_Voltaje(2, 2)

        'ComboBoxEdit9.SelectedIndex = 0
        'ComboBoxEdit9.Items.Add("[Unit]")
        carga_Capacity(2, 3)


        'ComboBoxEdit10.SelectedIndex = 0
        'ComboBoxEdit10.Items.Add("[Unit]")
        carga_Temp(2, 1)

        OilV_cb.SelectedIndex = 0
        OilV_cb.Items.Add("[Unit]")
        OilT_cb.SelectedIndex = 0
        OilT_cb.Items.Add("[Unit]")
    End Sub

    Private Sub load_combos(ByVal tipo As Integer, ByVal tipo2 As Integer)

        'Dim dt As New DataTable
        'Dim ds As New DataSet
        If conectado() Then
            Using cmd As New SqlCommand("Catalog_Read", cnn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim param1 As New SqlParameter("@Tipo", SqlDbType.Int)
                param1.Value = tipo
                Dim param2 As New SqlParameter("@Tipo2", SqlDbType.Int)
                param2.Value = tipo2
                cmd.Parameters.Add(param1)
                cmd.Parameters.Add(param2)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                Select Case tipo2
                    Case 1
                        Voltage_cb.Items.Clear()
                        While dr.Read()
                            Voltage_cb.Items.Add(dr.Item(0))
                        End While
                    Case 2
                        Capacity_cb.Items.Clear()
                        While dr.Read()

                            Capacity_cb.Items.Add(dr.Item(0))
                        End While
                    Case 3
                        WindTemp_cb.Items.Clear()
                        While dr.Read()
                            WindTemp_cb.Items.Add(dr.Item(0))
                        End While
                End Select

                'dt.Load(dr)
                'ds.Tables.Add(dt)
            End Using
        End If
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

    Private Sub carga_Voltaje(ByVal tipo As Integer, ByVal tipo2 As Integer)

        Voltage_cb.Items.Clear()

        Dim dt As New DataTable
        Dim ds As New DataSet
        conectado()
        Using cmd As New SqlCommand("Catalog_Read", cnn)
            cmd.CommandType = CommandType.StoredProcedure

            Dim param1 As New SqlParameter("@Tipo", SqlDbType.Int)
            param1.Value = tipo
            Dim param2 As New SqlParameter("@Tipo2", SqlDbType.Int)
            param2.Value = tipo2
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            While dr.Read()
                Voltage_cb.Items.Add(dr.Item(0))
            End While
            'dt.Load(dr)
            'ds.Tables.Add(dt)
        End Using
        desconectado()
    End Sub

    Private Sub carga_Capacity(ByVal tipo As Integer, ByVal tipo2 As Integer)

        Capacity_cb.Items.Clear()

        Dim dt As New DataTable
        Dim ds As New DataSet
        conectado()
        Using cmd As New SqlCommand("Catalog_Read", cnn)
            cmd.CommandType = CommandType.StoredProcedure

            Dim param1 As New SqlParameter("@Tipo", SqlDbType.Int)
            param1.Value = tipo
            Dim param2 As New SqlParameter("@Tipo2", SqlDbType.Int)
            param2.Value = tipo2
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            While dr.Read()
                Capacity_cb.Items.Add(dr.Item(0))
            End While
            'dt.Load(dr)
            'ds.Tables.Add(dt)
        End Using
        desconectado()
    End Sub

    Private Sub carga_Temp(ByVal tipo As Integer, ByVal tipo2 As Integer)

        WindTemp_cb.Items.Clear()

        Dim dt As New DataTable
        Dim ds As New DataSet
        conectado()
        Using cmd As New SqlCommand("Catalog_Read", cnn)
            cmd.CommandType = CommandType.StoredProcedure

            Dim param1 As New SqlParameter("@Tipo", SqlDbType.Int)
            param1.Value = tipo
            Dim param2 As New SqlParameter("@Tipo2", SqlDbType.Int)
            param2.Value = tipo2
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Dim dr As SqlDataReader = cmd.ExecuteReader()

            While dr.Read()
                WindTemp_cb.Items.Add(dr.Item(0))
            End While
            'dt.Load(dr)
            'ds.Tables.Add(dt)
        End Using
        desconectado()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Button1.Click

        Dim dd As Boolean = IIf(DETC_Yes.IsChecked, True, False)
        Dim ll As Boolean = IIf(LTC_Yes.IsChecked, True, False)
        Dim pp As Boolean = IIf(Pumps_Yes.IsChecked, True, False)


        Transformer_Insert(Id_txt.Text, SerNum_txt.Text, 0, Design_txt.Text, Manu_cb.SelectedIndex, Subs_cb.SelectedIndex, _
                           Utility_cb.SelectedIndex, Type_cb.SelectedIndex, Phases_cb.SelectedIndex, Core_cb.SelectedIndex, Voltage_txt.Text, _
                           Voltage_cb.SelectedIndex, Capacity_txt.Text, Capacity_cb.SelectedIndex, Year_txt.Text, Bank_txt.Text, Class_cb.SelectedIndex, _
                           OilV_txt.Text, OilV_cb.SelectedIndex, OilT_txt.Text, OilT_cb.SelectedIndex, SampleP_cb.SelectedIndex, _
                           dd, DETCP_txt.Text, ll, LTCM_txt.Text, 0, LTCT_txt.Text, _
                           LTCSP_txt.Text, LTCEP_txt.Text, CHV_txt.Text, HVConn_cb.SelectedIndex, CLV_txt.Text, _
                           LVConn_cb.SelectedIndex, WindTemp_txt.Text, WindTemp_cb.SelectedIndex, pp, NumP_txt.Text, NumG_txt.Text, Notes_txt.Text)


    End Sub

    Private Sub Transformer_Insert(ByVal Local_ID As Integer, ByVal Serial_Number As String, ByVal Design As Integer, ByVal Design_Desc As String, ByVal ID_Manufacturer As Integer, ByVal ID_Substation As Integer, _
                                   ByVal ID_Trafo_Utility As Integer, ByVal ID_Trafo_Type As Integer, ByVal ID_Trafo_Phases As Integer, ByVal ID_Core As Integer, ByVal Voltage As Decimal, _
                                   ByVal Volt_ID_Unit As Integer, ByVal Capacity As Decimal, ByVal Cap_ID_Unit As Integer, ByVal Year As Integer, ByVal Bank As String, ByVal ID_Trafo_Class As Integer, _
                                   ByVal Oil_Volume As Decimal, ByVal Oil_Vol_ID_Unit As Integer, ByVal Oil_Temp As Decimal, ByVal Oil_Temp_ID_Unit As Integer, ByVal ID_Samp_Point As Integer, _
                                   ByVal DETC As Boolean, ByVal DETC_Position As Integer, ByVal LTC As Boolean, ByVal LTC_Model As Integer, ByVal LTC_Operations As Integer, ByVal LTC_Type As Integer, _
                                   ByVal LTC_Position_Start As Integer, ByVal LTC_Position_End As Integer, ByVal HV_Current As Integer, ByVal ID_HV_Connection As Integer, ByVal LV_Current As Integer, _
                                   ByVal ID_LV_Connection As Integer, ByVal Winding_Temp As Decimal, ByVal Wind_Temp_ID_Unit As Integer, ByVal Pumps As Boolean, ByVal Number_Pumps As Integer, ByVal Number_Groups As Integer, _
                                   ByVal Notes As String)

        conectado()
        Using cmd As New SqlCommand("Transformer_Insert", cnn)
            cmd.CommandType = CommandType.StoredProcedure

            '----------
            Dim param1 As New SqlParameter("@Local_ID", SqlDbType.Int)
            param1.Value = Local_ID
            Dim param2 As New SqlParameter("@Serial_Number", SqlDbType.NVarChar)
            param2.Value = Serial_Number
            Dim param3 As New SqlParameter("@Design", SqlDbType.Int)
            param3.Value = Design
            Dim param4 As New SqlParameter("@Design_Desc", SqlDbType.NVarChar)
            param4.Value = Design_Desc
            Dim param5 As New SqlParameter("@ID_Manufacturer", SqlDbType.Int)
            param5.Value = ID_Manufacturer
            Dim param6 As New SqlParameter("@ID_Substation", SqlDbType.Int)
            param6.Value = ID_Substation
            Dim param7 As New SqlParameter("@ID_Trafo_Utility", SqlDbType.Int)
            param7.Value = ID_Trafo_Utility
            Dim param8 As New SqlParameter("@ID_Trafo_Type", SqlDbType.Int)
            param8.Value = ID_Trafo_Type
            Dim param9 As New SqlParameter("@ID_Trafo_Phases", SqlDbType.Int)
            param9.Value = ID_Trafo_Phases
            Dim param10 As New SqlParameter("@ID_Core", SqlDbType.Int)
            param10.Value = ID_Core
            Dim param11 As New SqlParameter("@Voltage", SqlDbType.Decimal)
            param11.Value = Voltage
            Dim param12 As New SqlParameter("@Volt_ID_Unit", SqlDbType.Int)
            param12.Value = Volt_ID_Unit
            Dim param13 As New SqlParameter("@Capacity", SqlDbType.Decimal)
            param13.Value = Capacity
            Dim param14 As New SqlParameter("@Cap_ID_Unit", SqlDbType.Int)
            param14.Value = Cap_ID_Unit
            Dim param15 As New SqlParameter("@Year", SqlDbType.Int)
            param15.Value = Year
            Dim param16 As New SqlParameter("@Bank", SqlDbType.NVarChar)
            param16.Value = Bank
            Dim param17 As New SqlParameter("@ID_Trafo_Class", SqlDbType.Int)
            param17.Value = ID_Trafo_Class
            Dim param18 As New SqlParameter("@Oil_Volume", SqlDbType.Decimal)
            param18.Value = Oil_Volume
            Dim param19 As New SqlParameter("@Oil_Vol_ID_Unit", SqlDbType.Int)
            param19.Value = Oil_Vol_ID_Unit
            Dim param20 As New SqlParameter("@Oil_Temp", SqlDbType.Decimal)
            param20.Value = Oil_Temp
            Dim param21 As New SqlParameter("@Oil_Temp_ID_Unit", SqlDbType.Int)
            param21.Value = Oil_Temp_ID_Unit
            Dim param22 As New SqlParameter("@ID_Samp_Point", SqlDbType.Int)
            param22.Value = ID_Samp_Point
            Dim param23 As New SqlParameter("@DETC", SqlDbType.Bit)
            param23.Value = DETC
            Dim param24 As New SqlParameter("@DETC_Position", SqlDbType.Int)
            param24.Value = DETC_Position
            Dim param25 As New SqlParameter("@LTC", SqlDbType.Bit)
            param25.Value = LTC
            Dim param26 As New SqlParameter("@LTC_Model", SqlDbType.Int)
            param26.Value = LTC_Model
            Dim param27 As New SqlParameter("@LTC_Operations", SqlDbType.Int)
            param27.Value = LTC_Operations
            Dim param28 As New SqlParameter("@LTC_Type", SqlDbType.Int)
            param28.Value = LTC_Type
            Dim param29 As New SqlParameter("@LTC_Position_Start", SqlDbType.Int)
            param29.Value = LTC_Position_Start
            Dim param30 As New SqlParameter("@LTC_Position_End", SqlDbType.Int)
            param30.Value = LTC_Position_End
            Dim param31 As New SqlParameter("@HV_Current", SqlDbType.Int)
            param31.Value = HV_Current
            Dim param32 As New SqlParameter("@ID_HV_Connection", SqlDbType.Int)
            param32.Value = ID_HV_Connection
            Dim param33 As New SqlParameter("@LV_Current", SqlDbType.Int)
            param33.Value = LV_Current
            Dim param34 As New SqlParameter("@ID_LV_Connection", SqlDbType.Int)
            param34.Value = ID_LV_Connection
            Dim param35 As New SqlParameter("@Winding_Temp", SqlDbType.Decimal)
            param35.Value = Winding_Temp
            Dim param36 As New SqlParameter("@Wind_Temp_ID_Unit", SqlDbType.Int)
            param36.Value = Wind_Temp_ID_Unit
            Dim param37 As New SqlParameter("@Pumps", SqlDbType.Bit)
            param37.Value = Pumps
            Dim param38 As New SqlParameter("@Number_Pumps", SqlDbType.Int)
            param38.Value = Number_Pumps
            Dim param39 As New SqlParameter("@Number_Groups", SqlDbType.Int)
            param39.Value = Number_Groups
            Dim param40 As New SqlParameter("@Notes", SqlDbType.NVarChar)
            param40.Value = Notes


            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param7)
            cmd.Parameters.Add(param8)
            cmd.Parameters.Add(param9)
            cmd.Parameters.Add(param10)
            cmd.Parameters.Add(param11)
            cmd.Parameters.Add(param12)
            cmd.Parameters.Add(param13)
            cmd.Parameters.Add(param14)
            cmd.Parameters.Add(param15)
            cmd.Parameters.Add(param16)
            cmd.Parameters.Add(param17)
            cmd.Parameters.Add(param18)
            cmd.Parameters.Add(param19)
            cmd.Parameters.Add(param20)
            cmd.Parameters.Add(param21)
            cmd.Parameters.Add(param22)
            cmd.Parameters.Add(param23)
            cmd.Parameters.Add(param24)
            cmd.Parameters.Add(param25)
            cmd.Parameters.Add(param26)
            cmd.Parameters.Add(param27)
            cmd.Parameters.Add(param28)
            cmd.Parameters.Add(param29)
            cmd.Parameters.Add(param30)
            cmd.Parameters.Add(param31)
            cmd.Parameters.Add(param32)
            cmd.Parameters.Add(param33)
            cmd.Parameters.Add(param34)
            cmd.Parameters.Add(param35)
            cmd.Parameters.Add(param36)
            cmd.Parameters.Add(param37)
            cmd.Parameters.Add(param38)
            cmd.Parameters.Add(param39)
            cmd.Parameters.Add(param40)
            If cmd.ExecuteNonQuery() Then
                MsgBox("Successful", MsgBoxStyle.Information, "Insert")
            Else
                MsgBox("Error!", MsgBoxStyle.Information, "Insert")
            End If
            desconectado()
        End Using
        desconectado()
    End Sub
End Class
