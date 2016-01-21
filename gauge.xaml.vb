Imports System.Drawing
Imports DevExpress.Xpf.Core
Imports DevExpress.Xpf.Gauges

Public Class gauge

    Private Sub gauge_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Dim m As New MenuP()
        m.Show()
    End Sub

    Private Sub gauge_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        CircularGaugeControl1.Scales(0).Needles(0).Value = 80
        CircularGaugeControl1.Scales(0).StartValue = 0
        CircularGaugeControl1.Scales(0).EndValue = 150
        CircularGaugeControl1.Scales(0).Ranges(0).StartValue = New RangeValue(0)
        CircularGaugeControl1.Scales(0).Ranges(0).EndValue = New RangeValue(50)
        CircularGaugeControl1.Scales(0).Ranges(1).StartValue = New RangeValue(50)
        CircularGaugeControl1.Scales(0).Ranges(1).EndValue = New RangeValue(100)
        CircularGaugeControl1.Scales(0).Ranges(2).StartValue = New RangeValue(100)
        CircularGaugeControl1.Scales(0).Ranges(2).EndValue = New RangeValue(150)
        CircularGaugeControl1.Scales(0).Markers(0).Value = 90
        CircularGaugeControl1.Scales(0).RangeBars(0).Value = 20
        CircularGaugeControl1.Scales(0).RangeBars(0).AnchorValue = 100

    End Sub
End Class
