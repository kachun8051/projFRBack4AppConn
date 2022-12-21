Imports System.ComponentModel
Imports System.Net
Imports Newtonsoft

Public Class clsFetchData

    Public Event DataDownloaded(isSuccess As Boolean)

    Public blRecord As BindingList(Of clsRecord)

    Private m_FromDate, m_ToDate, m_RangeType As String

    Public ReadOnly Property startDate() As String
        Get
            Return m_FromDate
        End Get
    End Property

    Public ReadOnly Property endDate() As String
        Get
            Return m_ToDate
        End Get
    End Property

    Public ReadOnly Property rangeType() As String
        Get
            Return m_RangeType
        End Get
    End Property

    Public Function getRangeRecords(fromdate As Date, todate As Date) As Boolean
        ' using TLS 1.2
        ' System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        Dim myurl As String = "https://parseapi.back4app.com/classes/Production?" &
            "where=" & getDateParam(fromdate, todate) & "&order=pickedAt"
        Diagnostics.Debug.WriteLine("Url: " & myurl)
        Dim web As New WebClient
        web.Headers.Add(HttpRequestHeader.Accept, "application/json")
        web.Headers.Add(HttpRequestHeader.ContentType, "application/json")
        web.Headers.Add("X-Parse-Application-Id", "5vUD5SzypdFDZfa7Sxjya1yLliHMAJ52ML3sqBf6")
        web.Headers.Add("X-Parse-REST-API-Key", "sgyDDR9YYlvTfkZv1datnUu75nhnnjqejm2yMFNL")
        web.Encoding = System.Text.Encoding.UTF8
        'Add the event handler here
        AddHandler web.DownloadStringCompleted, AddressOf webClient_DownloadStringCompleted2
        Try
            Dim objuri As Uri = New Uri(myurl)
            web.DownloadStringAsync(objuri)
            Return True
        Catch ex As Exception
            Console.WriteLine("clsRecords.getRangeRecords: " & vbCrLf & ex.Message)
            RaiseEvent DataDownloaded(False)
            Return False
        End Try
    End Function

    Private Sub webClient_DownloadStringCompleted2(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        Console.WriteLine(e.Result)
        Dim lstData As List(Of clsRecord) = Nothing
        Dim jsonResultToDict As Dictionary(Of String, List(Of clsRecord)) =
                Json.JsonConvert.DeserializeObject(Of Dictionary(Of String, List(Of clsRecord)))(e.Result)
        jsonResultToDict.TryGetValue("results", lstData)
        ConvertToBindingList(lstData)
        Console.WriteLine("Event is triggered")
        'objReport.ShowReport(blRecord)
        RaiseEvent DataDownloaded(True)
    End Sub

    Private Function getDateParam(i_from As Date, i_to As Date) As String

        m_FromDate = i_from.Year & "/" &
            i_from.Month.ToString.PadLeft(2, "0") & "/" &
            i_from.Day.ToString.PadLeft(2, "0")
        m_ToDate = i_to.Year & "/" &
            i_to.Month.ToString.PadLeft(2, "0") & "/" &
            i_to.Day.ToString.PadLeft(2, "0")
        If isTwoDatesSame(i_from, i_to) Then
            m_RangeType = "daily"
            Return getParam(i_from)
        End If
        m_RangeType = ""
        If isTwoDatesAWeek(i_from, i_to) Then
            m_RangeType = "weekly"
        End If
        If isTwoDatesAMonth(i_from, i_to) Then
            m_RangeType = "monthly"
        End If
        Return getParams(i_from, i_to)

    End Function

    Private Function getParams(i_date1 As Date, i_date2 As Date) As String
        Dim jobj_3F As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"__type", "Date"},
            {"iso", getISODate(i_date1)}
        }
        Dim jobj_3T As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"__type", "Date"},
            {"iso", getISODate(i_date2)}
        }
        Dim jobj_1 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"$gte", jobj_3F},
            {"$lt", jobj_3T}
        }
        Dim jobj_2 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"packedAt", jobj_1}
        }
        Dim strOutput As String = jobj_2.ToString(Newtonsoft.Json.Formatting.None)
        Diagnostics.Debug.WriteLine(strOutput)
        Return strOutput
    End Function

    Private Function getParam(i_date As Date) As String
        Dim jobj_1a As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"__type", "Date"},
            {"iso", getISODate(i_date)}
        }
        Dim jobj_1b As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"__type", "Date"},
            {"iso", getISODate(i_date.AddDays(1))}
        }
        Dim jobj_1 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"$gte", jobj_1a},
            {"$lt", jobj_1b}
        }
        Dim jobj_2 As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
            {"packedAt", jobj_1}
        }
        Dim strOutput As String = jobj_2.ToString(Json.Formatting.None)
        Diagnostics.Debug.WriteLine(strOutput)
        Return strOutput
    End Function

    Private Function getISODate(i_date As Date) As String
        Dim timeZoneHour As Int32 = getTimeZone()
        Dim strdt As String = i_date.Year.ToString.PadLeft(4, "0"c) & "-" &
            i_date.Month.ToString.PadLeft(2, "0"c) & "-" &
            i_date.Day.ToString.PadLeft(2, "0"c)
        Dim dt As Date = DateTime.Parse(strdt)
        dt = dt.AddHours(-timeZoneHour)
        Dim dt_1 As String = dt.ToString("yyyy-MM-dd HH:mm:ss").Replace(" ", "T") & "Z"
        Diagnostics.Debug.WriteLine("ISO Date: " & dt_1)
        Return dt_1
    End Function

    ' Return positive or negative number
    Private Function getTimeZone() As Int32
        Dim tzi As TimeZoneInfo = TimeZoneInfo.Local
        Dim offset As TimeSpan = tzi.BaseUtcOffset
        Return offset.Hours
    End Function

    Private Sub ConvertToBindingList(ByRef lst As List(Of clsRecord))
        If blRecord Is Nothing Then
            blRecord = New BindingList(Of clsRecord)
        Else
            blRecord.Clear()
        End If
        If lst Is Nothing Then
            Return
        End If
        For Each obj As clsRecord In lst
            blRecord.Add(obj)
        Next

    End Sub

    Private Function isTwoDatesSame(dt1 As DateTime, dt2 As DateTime) As Boolean
        If dt1.Day <> dt2.Day Then
            Return False
        End If
        If dt1.Month <> dt2.Month Then
            Return False
        End If
        If dt1.Year <> dt2.Year Then
            Return False
        End If
        Return True
    End Function

    Private Function isTwoDatesAMonth(dt1 As DateTime, dt2 As DateTime) As Boolean
        Dim date_1 As DateTime = New DateTime(dt1.Year, dt1.Month, dt1.Day, 0, 0, 0)
        Dim date_2 As DateTime = New DateTime(dt2.Year, dt2.Month, dt2.Day, 0, 0, 0)
        If date_1.AddMonths(1).AddDays(-1) = date_2 Then
            Return True
        End If
        Return False
    End Function

    Private Function isTwoDatesAWeek(dt1 As DateTime, dt2 As DateTime) As Boolean
        Dim date_1 As DateTime = New DateTime(dt1.Year, dt1.Month, dt1.Day, 0, 0, 0)
        Dim date_2 As DateTime = New DateTime(dt2.Year, dt2.Month, dt2.Day, 0, 0, 0)
        If date_1.AddDays(6) = date_2 Then
            Return True
        End If
        Return False
    End Function

End Class
