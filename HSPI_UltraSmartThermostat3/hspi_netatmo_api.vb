Imports System.Net
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Text

Public Class hspi_netatmo_api

  Private _access_token As String = String.Empty
  Private _refresh_token As String = String.Empty
  Private _scope As String = String.Empty
  Private _expires_in As Integer = 0
  Private _refreshed As New Stopwatch

  Private _querySuccess As ULong = 0
  Private _queryFailure As ULong = 0

  Public Sub New()

  End Sub

  Public Function QuerySuccessCount() As ULong
    Return _querySuccess
  End Function

  Public Function QueryFailureCount() As ULong
    Return _queryFailure
  End Function

#Region "Netatmo Authentiation"

  ''' <summary>
  ''' Determines if the API key is valid
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidAPIToken() As Boolean

    Try

      If gAPIClientId.Length = 0 Then
        Return False
      ElseIf gAPIClientSecret.Length = 0 Then
        Return False
      ElseIf NetatmoAPI.GetAccessToken.Length = 0 Then
        Return False
      Else
        Return True
      End If

    Catch pEx As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Get Access Token
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAccessToken() As String

    Dim expires_in As Long = _refreshed.ElapsedMilliseconds / 1000

    If CheckCredentials() = False Then
      WriteMessage("Invalid API authentication information.  Please check plug-in options.", MessageType.Error)
    ElseIf _access_token.Length = 0 Then
      GetToken()
    ElseIf expires_in > _expires_in Then
      _access_token = String.Empty
      RefreshAccessToken()
    End If

    Return _access_token

  End Function

  ''' <summary>
  ''' Gets Accesss Token
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetToken()

    Try

      Dim data As Byte() = New ASCIIEncoding().GetBytes(String.Format("grant_type={0}&client_id={1}&client_secret={2}&username={3}&password={4}&scope={5}", "password", gAPIClientId, gAPIClientSecret, gAPIUsername, gAPIPassword, gAPIScope))

      Dim strURL As String = String.Format("https://api.netatmo.com/oauth2/token")
      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)

      HTTPWebRequest.Timeout = 1000 * 60
      HTTPWebRequest.Method = "POST"
      HTTPWebRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
      HTTPWebRequest.ContentLength = data.Length

      Dim myStream As Stream = HTTPWebRequest.GetRequestStream
      myStream.Write(data, 0, data.Length)

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())

          Dim JSONString As String = reader.ReadToEnd()
          ' {"error":"invalid_client"}

          Dim js As New JavaScriptSerializer()
          Dim OAuth20 As oauth_token = js.Deserialize(Of oauth_token)(JSONString)

          If OAuth20.error.Length > 0 Then
            Throw New Exception(OAuth20.error)
          End If

          ' {"access_token":"53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984","refresh_token":"53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5","scope":["read_station"],"expires_in":10800}

          _access_token = OAuth20.access_token
          _refresh_token = OAuth20.refresh_token
          _expires_in = OAuth20.expires_in

          _refreshed.Start()

        End Using

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0

    End Try

  End Sub

  ''' <summary>
  ''' Checks to see if required credentials are available
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CheckCredentials() As Boolean

    Try

      Dim sbWarning As New StringBuilder

      If gAPIClientId.Length = 0 Then
        sbWarning.Append("Netatmo Client Id")
      End If
      If gAPIClientSecret.Length = 0 Then
        sbWarning.Append("Netatmo Client Secret")
      End If
      If gAPIUsername.Length = 0 Then
        sbWarning.Append("Netatmo Client Username")
      End If
      If gAPIPassword.Length = 0 Then
        sbWarning.Append("Netatmo Client Password")
      End If
      If sbWarning.Length = 0 Then Return True

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)
    End Try

    Return False

  End Function

  ''' <summary>
  ''' Refreshes the Access Token
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub RefreshAccessToken()

    Try

      Dim data As Byte() = New ASCIIEncoding().GetBytes(String.Format("grant_type={0}&client_id={1}&client_secret={2}&refresh_token={3}", "refresh_token", gAPIClientId, gAPIClientSecret, _refresh_token))

      Dim strURL As String = String.Format("https://api.netatmo.com/oauth2/token")
      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)

      HTTPWebRequest.Timeout = 1000 * 60
      HTTPWebRequest.Method = "POST"
      HTTPWebRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
      HTTPWebRequest.ContentLength = data.Length

      Dim myStream As Stream = HTTPWebRequest.GetRequestStream
      myStream.Write(data, 0, data.Length)

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())

          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          Dim OAuth20 As refresh_token = js.Deserialize(Of refresh_token)(JSONString)

          ' {"access_token":"53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984","refresh_token":"53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5","scope":["read_station"],"expires_in":10800}

          _access_token = OAuth20.access_token
          _refresh_token = OAuth20.refresh_token
          _expires_in = OAuth20.expires_in

          _refreshed.Start()

        End Using

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0

    End Try

  End Sub
#End Region

  ''' <summary>
  ''' Gets the Realtime Weather from UltraNetatmoThermostat
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetThermostatData() As ThermostatData

    Dim ThermostatData As New ThermostatData

    Try

      Dim access_token As String = Me.GetAccessToken()
      If access_token.Length = 0 Then
        Throw New Exception("Unable to get Netatmo Access Token.")
      End If

      Dim strURL As String = String.Format("https://api.netatmo.com/api/getthermostatsdata?access_token={0}", access_token)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          ThermostatData = js.Deserialize(Of ThermostatData)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Warning)

      _queryFailure += 1

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0
    End Try

    Return ThermostatData

  End Function

  ''' <summary>
  ''' Changes the Thermostat manual temperature setpoint
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="module_id"></param>
  ''' <param name="setpoint_mode"></param>
  ''' <param name="setpoint_endtime"></param>
  ''' <param name="setpoint_temp"></param>
  ''' <returns></returns>
  Public Function Setthermpoint(ByVal device_id As String,
                                ByVal module_id As String,
                                ByVal setpoint_mode As String,
                                Optional ByVal setpoint_endtime As Integer = -1,
                                Optional ByVal setpoint_temp As Double = -1) As SetthermpointResponse

    Dim SetthermpointResponse As New SetthermpointResponse

    Try

      Dim access_token As String = Me.GetAccessToken()
      If access_token.Length = 0 Then
        Throw New Exception("Unable to get Netatmo Access Token.")
      End If

      Dim strURL As String = String.Format("https://api.netatmo.com/api/setthermpoint?access_token={0}&device_id={1}&module_id={2}&setpoint_mode={3}", access_token, device_id, module_id, setpoint_mode)

      If setpoint_endtime >= 0 Then
        Select Case setpoint_mode
          Case "manual", "max"
            strURL &= String.Format("&setpoint_endtime={0}", setpoint_endtime)
        End Select
      End If

      If setpoint_temp >= 0 Then
        Select Case setpoint_mode
          Case "manual"
            strURL &= String.Format("&setpoint_temp={0}", setpoint_temp)
        End Select
      End If

      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          SetthermpointResponse = js.Deserialize(Of SetthermpointResponse)(JSONString)

        End Using

      End Using

      _querySuccess += 1

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(String.Format("An unexpected error was returned from the Setthermpoint Netatmo API call:  {0}", pEx.Message), MessageType.Error)

      _queryFailure += 1
    End Try

    Return SetthermpointResponse

  End Function

#Region "UltraNetatmoThermostat oAuth Token"

  '{
  '    "access_token": "53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984",
  '    "refresh_token": "53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5",
  '    "expires_in": 10800
  '}

  <Serializable()>
  Private Class oauth_token
    Public Property [error] As String = String.Empty
    Public Property access_token As String = String.Empty   ' Access token for your user
    Public Property refresh_token As String = String.Empty  ' Use this token to get a new access_token once it has expired
    Public Property expires_in As Integer = 0               ' Validity timelaps in seconds
  End Class

  <Serializable()>
  Private Class refresh_token
    Public Property access_token As String = String.Empty
    Public Property refresh_token As String = String.Empty  ' Refresh tokens do not change
    Public Property expires_in As Integer = 0
  End Class

#End Region

#Region "ThermostatData"

  <Serializable()>
  Public Class ThermostatData
    Public Property body As DevicesAndUser
    Public Property status As String = String.Empty
    Public Property time_exec As Double
    Public Property time_server As Double
  End Class

  <Serializable()>
  Public Class DevicesAndUser
    Public Property devices As List(Of Device)
    Public Property user As User
  End Class

  <Serializable()>
  Public Class Device
    Public Property _id As String = String.Empty
    Public Property firmware As Long = 0                      ' Version of the software
    Public Property last_setup As Long = 0
    Public Property last_status_store As Integer = 0
    Public Property modules As List(Of DeviceModule)          ' List of modules associated with the station and their details
    Public Property place As Place
    Public Property plug_connected_boiler As Long = 0
    Public Property station_name As String = String.Empty
    Public Property type As String = String.Empty
    Public Property udp_conn As Boolean = False
    Public Property wifi_status As Integer = 0
    Public Property last_plug_seen As Long = 0
  End Class

  <Serializable()>
  Public Class DeviceModule
    Public Property _id As String = String.Empty
    Public Property module_name As String = String.Empty
    Public Property type As String = String.Empty
    Public Property firmware As Long = 0
    Public Property last_message As Long = 0
    Public Property rf_status As Integer = 0                  ' Wifi status per Base station. (86=bad, 56=good) Find more details on the Weather Station page.
    Public Property battery_vp As Integer = 0                 ' Current battery status per module.
    Public Property therm_orientation As Long = 0
    Public Property therm_relay_cmd As Long = 0
    Public Property battery_percent As Integer = 0
    Public Property last_therm_seen As Long = 0
    Public Property setpoint As setpoint
    Public Property measured As measured
  End Class

  <Serializable()>
  Public Class Place
    Public Property altitude As Double = 0
    Public Property city As String = String.Empty
    Public Property country As String = String.Empty
    Public Property timezone As String = String.Empty
    Public Property location As Double()
  End Class

  <Serializable()>
  Public Class setpoint
    Public Property setpoint_temp As Double = 0.0
    Public Property setpoint_endtime As Long = 0
    Public Property setpoint_mode As String = String.Empty
  End Class

  <Serializable()>
  Public Class measured
    Public Property time As Long = 0
    Public Property temperature As Double = 0.0
    Public Property setpoint_temp As Double = 0
  End Class

  <Serializable()>
  Public Class SetthermpointResponse
    Public status As String = String.Empty
    Public time_exec As Double = 0.0
  End Class

  <Serializable()>
  Public Class User
    Public Property mail As String = String.Empty
    Public Property administrative As UserOption
  End Class

  <Serializable()>
  Public Class UserOption
    Public Property country As String = String.Empty
    Public Property reg_locale As String = String.Empty ' user regional preferences (used for displaying date)
    Public Property lang As String = String.Empty
    Public Property unit As Integer = 0                ' 0 -> metric system, 1 -> imperial system
    Public Property windunit As Integer = 0            ' 0 -> kph, 1 -> mph, 2 -> ms, 3 -> beaufort, 4 -> knot
    Public Property pressureunit As Integer = 0        ' 0 -> mbar, 1 -> inHg, 2 -> mmHg
    Public Property feel_like_algo As Integer = 0      ' 0 -> humidex, 1 -> heat-index 
  End Class

#End Region

End Class
