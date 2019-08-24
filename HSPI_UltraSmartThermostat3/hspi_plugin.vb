Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Text
Imports HomeSeerAPI
Imports Scheduler
Imports System.ComponentModel

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable

  Const Pagename = "Events"

  Public HSDevices As New SortedList

  Public NetatmoAPI As hspi_netatmo_api
  Public NetatmoAPILock As New Object

  Public gAPIClientId As String = "58877004e8ede183b28b92fa"
  Public gAPIClientSecret As String = "jo8doZsY96wyvQSW5VCZ6xDElC8GnrmbPCHRaVOy"
  Public gAPIUsername As String = String.Empty
  Public gAPIPassword As String = String.Empty
  Public gAPIScope As String = "read_thermostat write_thermostat"
  Public gAPIGetFavorites As Boolean = False

  Public gAPICultureInfo As String = "en-us"

  Public Const IFACE_NAME As String = "UltraSmartThermostat3"

  Public Const LINK_TARGET As String = "hspi_ultrasmartthermostat3/hspi_ultrasmartthermostat3.aspx"
  Public Const LINK_URL As String = "hspi_ultrasmartthermostat3.html"
  Public Const LINK_TEXT As String = "UltraSmartThermostat3"
  Public Const LINK_PAGE_TITLE As String = "UltraSmartThermostat3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultrasmartthermostat3/UltraSmartThermostat3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = String.Empty
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultrasmartthermostat3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public gRTCReceived As Boolean = False
  Public gDeviceValueType As String = "1"
  Public gDeviceImage As Boolean = True
  Public gStatusImageSizeWidth As Integer = 32
  Public gStatusImageSizeHeight As Integer = 32

  Public gTempUnit As Integer = 1
  Public gWindUnit As Integer = 1
  Public gRainUnit As Integer = 1
  Public gPressureUnit As Integer = 1
  Public gFeelsLike As Integer = 1

  Public gMonitoring As Boolean = True

#Region "UltraSmartThermostat3 Public Functions"

  ''' <summary>
  ''' Thermostat Data Thread
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub CheckThermostatDataThread()

    Dim strMessage As String = ""
    Dim iCheckInterval As Integer = 0

    Dim bAbortThread As Boolean = False

    Try
      '
      ' Begin Main While Loop
      '

      While bAbortThread = False

        Try

          ' Sleep for 5 seconds to allow the Netatmo API to "catch up"
          Thread.Sleep(5000)

          If gMonitoring = True Then
            SyncLock NetatmoAPILock
              CheckThermostatData()
            End SyncLock
          End If

          '
          ' Sleep the requested number of minutes between runs
          '
          iCheckInterval = CInt(hs.GetINISetting("Options", "CheckThermostatDataThread", "5", gINIFile))
          Thread.Sleep(1000 * (60 * iCheckInterval))

        Catch pEx As ThreadAbortException
          ' 
          ' There was a normal request to terminate the thread.  
          '
          bAbortThread = True      ' Not actually needed
          WriteMessage(String.Format("CheckThermostatDataThread received abort request, terminating normally."), MessageType.Debug)
        Catch pEx As Exception
          ' Ignore this exception
        End Try

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckThermostatDataThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckThermostatDataThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Checks Thermostat Data
  ''' </summary>
  Public Sub CheckThermostatData()

    Try

      Dim unitType As String = hs.GetINISetting("Options", "UnitType", "1", gINIFile)

      Dim ThermostatData As hspi_netatmo_api.ThermostatData = NetatmoAPI.GetThermostatData
      If Not IsNothing(ThermostatData.body) Then

        Dim User As hspi_netatmo_api.User = ThermostatData.body.user
        If Not IsNothing(User) Then
          Dim UserOption As hspi_netatmo_api.UserOption = ThermostatData.body.user.administrative

          gTempUnit = UserOption.unit
          gWindUnit = UserOption.windunit
          gRainUnit = gTempUnit
          gPressureUnit = UserOption.pressureunit
          gFeelsLike = UserOption.feel_like_algo

        Else

          Select Case unitType
            Case "0"
              gTempUnit = 1
              gWindUnit = 1
              gRainUnit = gTempUnit
              gPressureUnit = 1
              gFeelsLike = 1
            Case "1"
              gTempUnit = 0
              gWindUnit = 0
              gRainUnit = gTempUnit
              gPressureUnit = 0
              gFeelsLike = 0
          End Select

        End If

        For Each Device In ThermostatData.body.devices
          Dim dv_root_addr As String = Device._id
          Dim dv_root_type As String = Device.type
          Dim dv_root_name As String = String.Empty

          Select Case dv_root_type.ToLower
            Case "naplug"
              dv_root_type = "Netatmo Thermostat Relay"
          End Select

          dv_root_name = String.Format("{0} [{1}]", Device.station_name, dv_root_type)

          '
          ' Update the Device Firmware
          '
          Dim device_firmware As Integer = Device.firmware
          If device_firmware >= 0 Then
            Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Firmware")
            Dim dv_name As String = "Device Firmware Version"
            Dim dv_type As String = "Netatmo Firmware"

            GetHomeSeerDevice(Device._id, "", dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
            SetDeviceValue(dv_addr, device_firmware)
          End If

          '
          ' Update the WiFi Status
          '
          Dim WiFiStatus As Integer = Device.wifi_status
          If WiFiStatus > 0 Then
            Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "WiFi-Signal")
            Dim dv_name As String = "WiFi Signal Status"
            Dim dv_type As String = "Netatmo WiFi"

            GetHomeSeerDevice(Device._id, "", dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
            SetDeviceValue(dv_addr, WiFiStatus)
          End If

          '
          ' Update the Modules Count
          '
          Dim module_count As Integer = Device.modules.Count
          If module_count >= 0 Then
            Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Modules")
            Dim dv_name As String = "Module Count"
            Dim dv_type As String = "Netatmo Modules"

            GetHomeSeerDevice(Device._id, "", dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
            SetDeviceValue(dv_addr, module_count)
          End If

          '
          ' Update the Last Status Store Value
          '
          Dim time_utc As Long = Device.last_status_store
          If time_utc > 0 Then
            Dim time_now As Long = ConvertDateTimeToEpoch(DateTime.Now)
            Dim dv_value As Integer = time_now - time_utc
            dv_value = dv_value / 60

            Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "LastUpdate")
            Dim dv_name As String = "Last Update"
            Dim dv_type As String = "Netatmo Update"

            GetHomeSeerDevice(Device._id, "", dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
            SetDeviceValue(dv_addr, dv_value)
          End If

          '
          ' Process device modules
          '
          Dim Modules As List(Of hspi_netatmo_api.DeviceModule) = Device.modules
          For Each [Module] In Modules

            dv_root_addr = [Module]._id
            dv_root_type = [Module].type
            dv_root_name = String.Format("{0} [{1}]", Device.station_name, [Module].module_name)

            Select Case dv_root_type.ToLower
              Case "natherm1"
                dv_root_type = "Netatmo Thermostat Module"
            End Select

            '
            ' Update the Module Firmware
            '
            Dim module_firmware As Integer = [Module].firmware
            If module_firmware >= 0 Then
              Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Firmware")
              Dim dv_name As String = "Module Firmware Version"
              Dim dv_type As String = "Netatmo Firmware"

              GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
              SetDeviceValue(dv_addr, module_firmware)
            End If

            '
            ' Update the Battery Level
            '
            Dim Battery As Integer = [Module].battery_vp
            If Battery > 0 Then
              Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Battery-Status")
              Dim dv_name As String = "Battery Status"
              Dim dv_type As String = "Netatmo Battery"

              GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
              SetDeviceValue(dv_addr, Battery)
            End If

            '
            ' Update the Battery Percentage
            '
            Dim battery_percent As Integer = [Module].battery_percent
            If battery_percent >= 0 Then
              Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Battery-Percentage")
              Dim dv_name As String = "Battery Percentage"
              Dim dv_type As String = "Netatmo Battery"

              GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
              SetDeviceValue(dv_addr, battery_percent)
            End If

            '
            ' Update the rf_status
            '
            Dim RFStatus As Integer = [Module].rf_status
            If RFStatus > 0 Then
              Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "RF-Signal")
              Dim dv_name As String = "RF Signal Status"
              Dim dv_type As String = "Netatmo RF"

              GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
              SetDeviceValue(dv_addr, RFStatus)
            End If

            '
            ' Update the Last Measure Time
            '
            Dim measured_time_utc As Long = [Module].measured.time
            If measured_time_utc > 0 Then
              Dim time_now As Long = ConvertDateTimeToEpoch(DateTime.Now)
              Dim dv_value As Integer = time_now - measured_time_utc
              dv_value = dv_value / 60

              Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "LastUpdate")
              Dim dv_name As String = "Last Update"
              Dim dv_type As String = "Netatmo Update"

              GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
              SetDeviceValue(dv_addr, dv_value)
            End If

            '
            ' Update the Last Measure Temperature
            '
            Dim Temperature As Double = [Module].measured.temperature
            If Temperature > -999 Then
              If gTempUnit = 1 Then Temperature = (Temperature * 1.8) + 32

              Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Temperature")
              Dim dv_name As String = "Temperature"
              Dim dv_type As String = "Netatmo Temperature"

              GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
              SetDeviceValue(dv_addr, Temperature)
            End If

            Try
              '
              ' Update the Last Measure Setpoint Temperature
              '
              Dim Setpoint As Double = [Module].measured.setpoint_temp
              If Setpoint > -999 Then
                If gTempUnit = 1 Then Setpoint = (Setpoint * 1.8) + 32

                Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Setpoint")
                Dim dv_name As String = "Setpoint Temperature"
                Dim dv_type As String = "Netatmo Setpoint"

                GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
                SetDeviceValue(dv_addr, Setpoint)
              End If
            Catch pEx As Exception

            End Try

            Try
              Dim SetpointMode As String = [Module].setpoint.setpoint_mode
              If SetpointMode.Length > 0 Then
                Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Setpoint-Mode")
                Dim dv_name As String = "Setpoint Mode"
                Dim dv_type As String = "Netatmo Setpoint Mode"

                Dim SetpointModeValue As Integer = 0
                Select Case SetpointMode.ToLower
                  Case "off"
                    SetpointModeValue = ThermostatMode.off
                  Case "program"
                    SetpointModeValue = ThermostatMode.program
                  Case "away"
                    SetpointModeValue = ThermostatMode.away
                  Case "hg"
                    SetpointModeValue = ThermostatMode.hg
                  Case "manual"
                    SetpointModeValue = ThermostatMode.manual
                  Case "max"
                    SetpointModeValue = ThermostatMode.max
                End Select

                GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
                SetDeviceValue(dv_addr, SetpointModeValue)
              End If
            Catch pEx As Exception
              ' Call ProcessError(pEx, "CheckThermostatData()")
            End Try

            Try
              Dim setpoint_endtime_utc As Long = [Module].setpoint.setpoint_endtime
              If setpoint_endtime_utc >= 0 Then
                Dim time_now As Long = ConvertDateTimeToEpoch(DateTime.Now)
                Dim dv_value As Integer = setpoint_endtime_utc - time_now
                dv_value = dv_value / 60
                If dv_value < 0 Then dv_value = 0

                Dim dv_addr As String = String.Format("{0}-{1}", dv_root_addr, "Setpoint-Time")
                Dim dv_name As String = "Setpoint End Time"
                Dim dv_type As String = "Netatmo Setpoint Time"

                GetHomeSeerDevice(Device._id, [Module]._id, dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
                SetDeviceValue(dv_addr, dv_value)
              End If

            Catch pEx As Exception
              ' Call ProcessError(pEx, "CheckThermostatData()")
            End Try

          Next

        Next

      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckThermostatData()")
    End Try

  End Sub

  ''' <summary>
  ''' Returns the wind value
  ''' </summary>
  ''' <param name="strWindDirection"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWindDirectionValue(ByVal strWindDirection As String) As Double

    Try

      Dim windDirection As Integer = Val(strWindDirection)

      Select Case windDirection
        Case 0 To 11 : Return 0
        Case 12 To 34 : Return 22.5
        Case 35 To 56 : Return 45
        Case 57 To 78 : Return 67.5
        Case 79 To 101 : Return 90
        Case 102 To 123 : Return 112.5
        Case 123 To 146 : Return 135
        Case 147 To 168 : Return 157.5
        Case 169 To 191 : Return 180
        Case 192 To 213 : Return 202.5
        Case 214 To 236 : Return 225
        Case 237 To 258 : Return 247.5
        Case 259 To 291 : Return 270
        Case 282 To 303 : Return 292.5
        Case 204 To 236 : Return 315
        Case 237 To 348 : Return 337.5
        Case 349 To 359 : Return 0
        Case Else : Return -1
      End Select

    Catch pEx As Exception
      Return -1
    End Try

  End Function

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String,
                             ByVal strKey As String,
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered GetSetting() function.", MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to decrypt the data
      '
      If strSection = "API" And strKey = "Password" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      End If

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  '''  Saves plug-in settings to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String,
                         ByVal strKey As String,
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered SaveSetting() subroutine.", MessageType.Debug)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      If strSection = "API" And strKey = "GetFavorites" Then
        gAPIGetFavorites = CBool(strValue)
      End If

      '
      ' Apply the API Consumer Key
      '
      If strSection = "API" And strKey = "ClientId" Then
        gAPIClientId = strValue
      End If

      '
      ' Apply the API Secret Key
      '
      If strSection = "API" And strKey = "ClientSecret" Then
        If strValue.Length = 0 Then Exit Sub
        gAPIClientSecret = strValue
      End If

      '
      ' Apply the API Username
      '
      If strSection = "API" And strKey = "Username" Then
        gAPIUsername = strValue
      End If

      '
      ' Apply the API Password
      '
      If strSection = "API" And strKey = "Password" Then
        If strValue.Length = 0 Then Exit Sub
        gAPIPassword = strValue
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "UltraSmartThermostat3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      'triggers.Add(o, "Weather Alert")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      'Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber
      Case WeatherTriggers.WeatherAlert
        Dim triggerName As String = GetEnumName(WeatherTriggers.WeatherAlert)

        '
        ' Start Alert Type
        '
        Dim ActionSelected As String = trigger.Item("AlertType")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "AlertType", UID, sUnique)

        Dim jqAlertType As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqAlertType.autoPostBack = True

        jqAlertType.AddItem("(Select Alert Type)", "", (ActionSelected = ""))
        Dim names As String() = System.Enum.GetNames(GetType(WeatherAlertTypes))
        For i As Integer = 0 To names.Length - 1
          Dim strOptionName = names(i)
          Dim strOptionValue = names(i)
          jqAlertType.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqAlertType.Build)

        '
        ' Start Station Name
        '
        ActionSelected = trigger.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", triggerName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("reported by")
        stb.Append(jqStation.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case WeatherTriggers.WeatherAlert
          Dim triggerName As String = GetEnumName(WeatherTriggers.WeatherAlert)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "AlertType_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("AlertType") = ActionValue

              Case InStr(sKey, triggerName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Station") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case WeatherTriggers.WeatherAlert
          If trigger.Item("AlertType") = "" Then Configured = False
          If trigger.Item("Station") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case WeatherTriggers.WeatherAlert
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = "Weather Alert"
          Dim strAlertType As String = trigger.Item("AlertType")

          Dim strStationNumber As String = trigger.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} on {2}", strAlertType, strTriggerName, strStationNumber)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Private Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case WeatherTriggers.WeatherAlert
                Dim strTriggerName As String = "Weather Alert Trigger"
                Dim strAlertType As String = trigger.Item("AlertType")
                Dim strStationNumber As String = trigger.Item("Station")

                Dim strTriggerCheck As String = String.Format("{0},{1},{2}", strTriggerName, strStationNumber, strAlertType)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      'actions.Add(o, "Email Notification")          ' 1
      'actions.Add(o, "Speak Weather")               ' 2
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = ActInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case WeatherActions.EmailNotification
        Dim actionName As String = GetEnumName(WeatherActions.EmailNotification)

        '
        ' Start EmailNotification
        '
        Dim ActionSelected As String = action.Item("Notification")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "Notification", UID, sUnique)

        Dim jqNotifications As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNotifications.autoPostBack = True

        jqNotifications.AddItem("(Select E-mail Notification)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Weather Conditions", "Weather Forecast", "Weather Alerts"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqNotifications.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqNotifications.Build)

        '
        ' Start Station Name
        '
        ActionSelected = action.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("from")
        stb.Append(jqStation.Build)

      Case WeatherActions.SpeakWeather
        Dim actionName As String = GetEnumName(WeatherActions.SpeakWeather)

        '
        ' Start Speak Weather
        '
        Dim ActionSelected As String = action.Item("Notification")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "Notification", UID, sUnique)

        Dim jqNotifications As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNotifications.autoPostBack = True

        jqNotifications.AddItem("(Select Speak Action)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Weather Conditions", "Weather Forecast", "Weather Alerts"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqNotifications.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqNotifications.Build)

        '
        ' Start Station Name
        '
        ActionSelected = action.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("from")
        stb.Append(jqStation.Build)

        '
        ' Start Speaker Host
        '
        ActionSelected = IIf(action.Item("SpeakerHost").Length = 0, "*:*", action.Item("SpeakerHost"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "SpeakerHost", UID, sUnique)

        Dim jqSpeakerHost As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 45, True)
        stb.Append("Host(host:instance)")
        stb.Append(jqSpeakerHost.Build)

    End Select

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case WeatherActions.EmailNotification
          Dim actionName As String = GetEnumName(WeatherActions.EmailNotification)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "Notification_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Notification") = ActionValue

              Case InStr(sKey, actionName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Station") = ActionValue

            End Select
          Next

        Case WeatherActions.SpeakWeather
          Dim actionName As String = GetEnumName(WeatherActions.SpeakWeather)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "Notification_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Notification") = ActionValue

              Case InStr(sKey, actionName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Station") = ActionValue

              Case InStr(sKey, actionName & "SpeakerHost_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("SpeakerHost") = ActionValue
            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case WeatherActions.EmailNotification
        If action.Item("Notification") = "" Then Configured = False
        If action.Item("Station") = "" Then Configured = False

      Case WeatherActions.SpeakWeather
        If action.Item("Notification") = "" Then Configured = False
        If action.Item("Station") = "" Then Configured = False
        If action.Item("SpeakerHost") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber

      Case WeatherActions.EmailNotification
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(WeatherActions.EmailNotification)

          Dim strNotificationType As String = action.Item("Notification")

          Dim strStationNumber As String = action.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} {2}", strActionName, strStationNumber, strNotificationType)
        End If

      Case WeatherActions.SpeakWeather
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(WeatherActions.SpeakWeather)

          Dim strNotificationType As String = action.Item("Notification")

          Dim strStationNumber As String = action.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          Dim strSpeakerHost As String = action.Item("SpeakerHost")

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} from {2} on {3}", strActionName, strNotificationType, strStationNumber, strSpeakerHost)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber

        Case WeatherActions.EmailNotification
          Dim strNotificationType As String = action.Item("Notification")
          Dim strStationNumber As String = action.Item("Station")

          Select Case strNotificationType
            Case "Weather Conditions"
              'EmailWeatherConditions(strStationNumber)

            Case "Weather Forecast"
              'EmailWeatherForecast(strStationNumber)

            Case "Weather Alerts"
              'EmailWeatherAlerts(strStationNumber)

          End Select

        Case WeatherActions.SpeakWeather
          Dim strNotificationType As String = action.Item("Notification")
          Dim strStationNumber As String = action.Item("Station")
          Dim strSpeakerHost As String = action.Item("SpeakerHost")

          Select Case strNotificationType
            Case "Weather Conditions"
              'SpeakWeatherConditions(strStationNumber, False, strSpeakerHost)

            Case "Weather Forecast"
              'SpeakWeatherForecast(strStationNumber, False, strSpeakerHost)

            Case "Weather Alerts"
              'SpeakWeatherAlerts(strStationNumber, False, strSpeakerHost)
          End Select

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum WeatherTriggers
  <Description("Weather Alert")> _
  WeatherAlert = 1
End Enum

Public Enum WeatherActions
  <Description("Email Notification")> _
  EmailNotification = 1
  <Description("Speak Weather")> _
  SpeakWeather = 2
End Enum

<Flags()> Public Enum WeatherAlertTypes
  Any = 0
  Forecast = 2
  Statement = 4
  Synopsis = 8
  Outlook = 16
  Watch = 32
  Advisory = 64
  Warning = 128
End Enum

<Flags()> Public Enum ThermostatMode
  off = 0
  program = 1
  away = 2
  hg = 3
  manual = 4
  max = 5
End Enum
