Imports System.ComponentModel
Imports System.IO
Imports System.Xml

Public Class clsReport
    Private ds As DataSet
    Private xmlFile, xsdFile, rptFile As String
    ' 3 parameters for report
    Private pStart, pEnd, pRangeType As String

    Sub New()
        xmlFile = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fastreport", "itemlist.xml")
        xsdFile = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fastreport", "itemlist.xsd")
        rptFile = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fastreport", "ProductionReport.frx")
    End Sub

    Public Sub setParams(dtStart As String, dtEnd As String, rType As String)
        pStart = dtStart
        pEnd = dtEnd
        pRangeType = rType
    End Sub

    Public Function ShowReport(ByRef blRecord As BindingList(Of clsRecord)) As Boolean
        Dim isInited As Boolean = InitTheDataSet()
        If isInited = False Then
            Return False
        End If
        Dim isFilled As Boolean = FillTheDataSet(blRecord)
        If isFilled = False Then
            Return False
        End If
        Try
            Dim rpt As New FastReport.Report
            rpt.Load(rptFile)
            If pStart <> String.Empty Then
                rpt.SetParameterValue("WhichStartDate", pStart)
            End If
            If pEnd <> String.Empty Then
                rpt.SetParameterValue("WhichEndDate", pEnd)
            End If
            If pRangeType <> String.Empty Then
                rpt.SetParameterValue("RangeType", pRangeType)
            End If
            rpt.RegisterData(ds)
            rpt.Show(vbFalse)
            Return True
        Catch ex As Exception
            Diagnostics.Debug.WriteLine("clsReport.ShowReport: " & vbCrLf & ex.Message)
            Return False
        End Try

    End Function

    Private Function InitTheDataSet() As Boolean
        Try
            Dim strXmlContent As String = System.IO.File.ReadAllText(xmlFile)
            Dim sr As StringReader = New StringReader(strXmlContent)
            Dim xtr As XmlTextReader = New XmlTextReader(sr)
            ds = New DataSet
            ds.ReadXml(xtr)
            ds.ReadXmlSchema(xsdFile)
            Return True
        Catch ex As Exception
            Diagnostics.Debug.WriteLine("clsReport.InitTheDataSet: " & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Function FillTheDataSet(ByRef lst As BindingList(Of clsRecord)) As Boolean
        If lst Is Nothing Then
            Return False
        End If
        If ds IsNot Nothing Then
            ds.Clear()
        End If
        For Each obj As clsRecord In lst
            Dim _row As DataRow = ds.Tables("item").NewRow()
            _row("objectId") = obj.objectId
            _row("itemnum") = obj.itemnum
            _row("itemname") = obj.itemname
            _row("itemname2") = obj.itemname2
            _row("itemuom") = obj.itemuom
            _row("itemstandardweight") = obj.itemstandardweight
            _row("itemprice") = obj.itemprice
            _row("weightingram") = obj.weightingram
            _row("sellingprice") = obj.sellingprice
            _row("barcode") = obj.barcode
            _row("packingdt") = obj.packingdt
            ds.Tables("item").Rows.Add(_row)
        Next
        Return True
    End Function

End Class
