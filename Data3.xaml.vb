Public Class Data3

    Private Sub Data3_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Dim m As New MenuP()
        m.Show()
    End Sub

    Private Sub DateEdit1_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles DateEdit1.EditValueChanged
        LblDate.Content = DateEdit1.Text.Substring(0, DateEdit1.Text.IndexOf(" "))
    End Sub

    Private Sub H2_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles H2_txt.EditValueChanged
        LblH2.Content = H2_txt.Text
    End Sub

    Private Sub CO_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles CO_txt.EditValueChanged
        LblCO.Content = CO_txt.Text
    End Sub

    Private Sub CO2_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles CO2_txt.EditValueChanged
        LblCO2.Content = CO2_txt.Text
    End Sub

    Private Sub CH4_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles CH4_txt.EditValueChanged
        LblCH4.Content = CH4_txt.Text
    End Sub

    Private Sub C2H6_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles C2H6_txt.EditValueChanged
        LblC2H6.Content = C2H6_txt.Text
    End Sub

    Private Sub C2H4_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles C2H4_txt.EditValueChanged
        LblC2H4.Content = C2H4_txt.Text
    End Sub

    Private Sub C2H2_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles C2H2_txt.EditValueChanged
        LblC2H2.Content = C2H2_txt.Text
    End Sub

    Private Sub H2_txt_LostFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles H2_txt.LostFocus
        If H2_txt.Text = "" Or Not IsNumeric(H2_txt.Text) Then
            H2_txt.Text = "0"
        End If

        Select Case Integer.Parse(H2_txt.Text)
            Case Is < 101
                H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#47d147"), SolidColorBrush)
                H2_Cond.Content = 1
            Case 101 To 700
                H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ffff33"), SolidColorBrush)
                H2_Cond.Content = 2
            Case 701 To 1800
                H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff8000"), SolidColorBrush)
                H2_Cond.Content = 3
            Case Is > 1800
                H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff0000"), SolidColorBrush)
                H2_Cond.Content = 4
        End Select

        LblTDCG.Content = Integer.Parse(CO_txt.Text) + Integer.Parse(H2_txt.Text) + Integer.Parse(CH4_txt.Text) + Integer.Parse(C2H6_txt.Text) _
            + Integer.Parse(C2H4_txt.Text) + Integer.Parse(C2H2_txt.Text)
        TDCG_txt.Text = LblTDCG.Content

    End Sub

    Private Sub CO_txt_LostFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles CO_txt.LostFocus
        If CO_txt.Text = "" Or Not IsNumeric(CO_txt.Text) Then
            CO_txt.Text = "0"
        End If

        Select Case Integer.Parse(CO_txt.Text)
            Case Is < 351
                CO_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#47d147"), SolidColorBrush)
                CO_Cond.Content = 1
            Case 351 To 571
                CO_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ffff33"), SolidColorBrush)
                CO_Cond.Content = 2
            Case 571 To 1400
                CO_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff8000"), SolidColorBrush)
                CO_Cond.Content = 3
            Case Is > 1400
                CO_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff0000"), SolidColorBrush)
                CO_Cond.Content = 4
        End Select

        LblTDCG.Content = Integer.Parse(CO_txt.Text) + Integer.Parse(H2_txt.Text) + Integer.Parse(CH4_txt.Text) + Integer.Parse(C2H6_txt.Text) _
            + Integer.Parse(C2H4_txt.Text) + Integer.Parse(C2H2_txt.Text)
        TDCG_txt.Text = LblTDCG.Content
    End Sub

    Private Sub CO2_txt_LostFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles CO2_txt.LostFocus
        If CO2_txt.Text = "" Or Not IsNumeric(CO2_txt.Text) Then
            CO2_txt.Text = "0"
        End If

        Select Case Integer.Parse(CO2_txt.Text)
            Case Is < 2501
                CO2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#47d147"), SolidColorBrush)
                CO2_Cond.Content = 1
            Case 2501 To 4001
                CO2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ffff33"), SolidColorBrush)
                CO2_Cond.Content = 2
            Case 4001 To 10000
                CO2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff8000"), SolidColorBrush)
                CO2_Cond.Content = 3
            Case Is > 10000
                CO2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff0000"), SolidColorBrush)
                CO2_Cond.Content = 4
        End Select

        LblTDCG.Content = Integer.Parse(CO_txt.Text) + Integer.Parse(H2_txt.Text) + Integer.Parse(CH4_txt.Text) + Integer.Parse(C2H6_txt.Text) _
            + Integer.Parse(C2H4_txt.Text) + Integer.Parse(C2H2_txt.Text)
        TDCG_txt.Text = LblTDCG.Content
    End Sub

    Private Sub CH4_txt_LostFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles CH4_txt.LostFocus
        If CH4_txt.Text = "" Or Not IsNumeric(CH4_txt.Text) Then
            CH4_txt.Text = "0"
        End If

        Select Case Integer.Parse(CH4_txt.Text)
            Case Is < 121
                CH4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#47d147"), SolidColorBrush)
                CH4_Cond.Content = 1
            Case 121 To 401
                CH4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ffff33"), SolidColorBrush)
                CH4_Cond.Content = 2
            Case 401 To 1000
                CH4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff8000"), SolidColorBrush)
                CH4_Cond.Content = 3
            Case Is > 1000
                CH4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff0000"), SolidColorBrush)
                CH4_Cond.Content = 4
        End Select

        LblTDCG.Content = Integer.Parse(CO_txt.Text) + Integer.Parse(H2_txt.Text) + Integer.Parse(CH4_txt.Text) + Integer.Parse(C2H6_txt.Text) _
            + Integer.Parse(C2H4_txt.Text) + Integer.Parse(C2H2_txt.Text)
        TDCG_txt.Text = LblTDCG.Content
    End Sub

    Private Sub C2H6_txt_LostFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles C2H6_txt.LostFocus
        If C2H6_txt.Text = "" Or Not IsNumeric(C2H6_txt.Text) Then
            C2H6_txt.Text = "0"
        End If

        Select Case Integer.Parse(C2H6_txt.Text)
            Case Is < 66
                C2H6_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#47d147"), SolidColorBrush)
                C2H6_Cond.Content = 1
            Case 66 To 101
                C2H6_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ffff33"), SolidColorBrush)
                C2H6_Cond.Content = 2
            Case 101 To 150
                C2H6_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff8000"), SolidColorBrush)
                C2H6_Cond.Content = 3
            Case Is > 150
                C2H6_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff0000"), SolidColorBrush)
                C2H6_Cond.Content = 4
        End Select

        LblTDCG.Content = Integer.Parse(CO_txt.Text) + Integer.Parse(H2_txt.Text) + Integer.Parse(CH4_txt.Text) + Integer.Parse(C2H6_txt.Text) _
            + Integer.Parse(C2H4_txt.Text) + Integer.Parse(C2H2_txt.Text)
        TDCG_txt.Text = LblTDCG.Content
    End Sub

    Private Sub C2H4_txt_LostFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles C2H4_txt.LostFocus
        If C2H4_txt.Text = "" Or Not IsNumeric(C2H4_txt.Text) Then
            C2H4_txt.Text = "0"
        End If

        Select Case Integer.Parse(C2H4_txt.Text)
            Case Is < 51
                C2H4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#47d147"), SolidColorBrush)
                C2H4_Cond.Content = 1
            Case 51 To 101
                C2H4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ffff33"), SolidColorBrush)
                C2H4_Cond.Content = 2
            Case 101 To 200
                C2H4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff8000"), SolidColorBrush)
                C2H4_Cond.Content = 3
            Case Is > 200
                C2H4_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff0000"), SolidColorBrush)
                C2H4_Cond.Content = 4
        End Select

        LblTDCG.Content = Integer.Parse(CO_txt.Text) + Integer.Parse(H2_txt.Text) + Integer.Parse(CH4_txt.Text) + Integer.Parse(C2H6_txt.Text) _
            + Integer.Parse(C2H4_txt.Text) + Integer.Parse(C2H2_txt.Text)
        TDCG_txt.Text = LblTDCG.Content
    End Sub

    Private Sub C2H2_txt_LostFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles C2H2_txt.LostFocus
        If C2H2_txt.Text = "" Or Not IsNumeric(C2H2_txt.Text) Then
            C2H2_txt.Text = "0"
        End If

        Select Case Integer.Parse(C2H2_txt.Text)
            Case Is < 2
                C2H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#47d147"), SolidColorBrush)
                C2H2_Cond.Content = 1
            Case 2 To 10
                C2H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ffff33"), SolidColorBrush)
                C2H2_Cond.Content = 2
            Case 10 To 35
                C2H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff8000"), SolidColorBrush)
                C2H2_Cond.Content = 3
            Case Is > 35
                C2H2_txt.Background = DirectCast(New BrushConverter().ConvertFromString("#ff0000"), SolidColorBrush)
                C2H2_Cond.Content = 4
        End Select

        LblTDCG.Content = Integer.Parse(CO_txt.Text) + Integer.Parse(H2_txt.Text) + Integer.Parse(CH4_txt.Text) + Integer.Parse(C2H6_txt.Text) _
            + Integer.Parse(C2H4_txt.Text) + Integer.Parse(C2H2_txt.Text)
        TDCG_txt.Text = LblTDCG.Content
    End Sub

    Private Sub TDCG_txt_EditValueChanged(sender As Object, e As DevExpress.Xpf.Editors.EditValueChangedEventArgs) Handles TDCG_txt.EditValueChanged
        H2_P.Content = Math.Round(((Integer.Parse(LblH2.Content) * 100) / Integer.Parse(LblTDCG.Content)), 2, MidpointRounding.ToEven)
        CO_P.Content = Math.Round(((Integer.Parse(LblCO.Content) * 100) / Integer.Parse(LblTDCG.Content)), 2, MidpointRounding.ToEven)
        'CO2_P.Content = Math.Round(((Integer.Parse(LblCO2.Content) * 100) / Integer.Parse(LblTDCG.Content)), 2, MidpointRounding.ToEven)
        CH4_P.Content = Math.Round(((Integer.Parse(LblCH4.Content) * 100) / Integer.Parse(LblTDCG.Content)), 2, MidpointRounding.ToEven)
        C2H6_P.Content = Math.Round(((Integer.Parse(LblC2H6.Content) * 100) / Integer.Parse(LblTDCG.Content)), 2, MidpointRounding.ToEven)
        C2H4_P.Content = Math.Round(((Integer.Parse(LblC2H4.Content) * 100) / Integer.Parse(LblTDCG.Content)), 2, MidpointRounding.ToEven)
        C2H2_P.Content = Math.Round(((Integer.Parse(LblC2H2.Content) * 100) / Integer.Parse(LblTDCG.Content)), 2, MidpointRounding.ToEven)
        TDCG_P.Content = Math.Round(H2_P.Content + CO_P.Content + CO2_P.Content + CH4_P.Content + C2H6_P.Content + C2H4_P.Content + C2H2_P.Content, 0, MidpointRounding.ToEven)

        Select Case Integer.Parse(TDCG_txt.Text)
            Case Is < 721
                TDCG_Cond.Content = 1
            Case 721 To 1920
                TDCG_Cond.Content = 2
            Case 1921 To 4631
                TDCG_Cond.Content = 3
            Case Is > 4631
                TDCG_Cond.Content = 4
        End Select
    End Sub

    Private Sub TDCG_txt_GotFocus(sender As Object, e As System.Windows.RoutedEventArgs) Handles TDCG_txt.GotFocus
        H2_txt.Focus()
    End Sub

End Class
