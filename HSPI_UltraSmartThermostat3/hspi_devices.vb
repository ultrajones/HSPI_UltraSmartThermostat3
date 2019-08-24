﻿Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service
Imports System.Text.RegularExpressions

Module hspi_devices

  Public DEV_INTERFACE As Byte = 1

  Dim bCreateRootDevice = True

  ''' <summary>
  ''' Function to initilize our plug-ins devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitPluginDevices() As String

    Dim strMessage As String = ""

    WriteMessage("Entered InitPluginDevices() function.", MessageType.Debug)

    Try
      Dim Devices As Byte() = {DEV_INTERFACE}
      For Each dev_cod As Byte In Devices
        Dim strResult As String = CreatePluginDevice(IFACE_NAME, dev_cod)
        If strResult.Length > 0 Then Return strResult
      Next

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "InitPluginDevices()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Subroutine to create a HomeSeer device
  ''' </summary>
  ''' <param name="base_code"></param>
  ''' <param name="dev_code"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreatePluginDevice(ByVal base_code As String, ByVal dev_code As String) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim dv_ref As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DevicePairs As New ArrayList
    Dim DeviceShowValues As Boolean = False

    Try

      Select Case dev_code
        Case DEV_INTERFACE.ToString
          '
          ' Create the UltraNetatmoThermostat State device
          '
          dv_name = "Netatmo State"
          dv_type = "Netatmo State"
          dv_addr = String.Concat(base_code, "-State")

        Case Else
          Throw New Exception(String.Format("Unable to create plug-in device for unsupported device name: {0}", dv_name))
      End Select

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = "Plug-ins"
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      If bCreateRootDevice = True Then
        dv.AssociatedDevice_ClearAll(hs)
        Dim dvp_ref As Integer = CreateRootDevice("Plugin", IFACE_NAME, IFACE_NAME, IFACE_NAME, dv_ref)
        If dvp_ref > 0 Then
          dv.AssociatedDevice_Add(hs, dvp_ref)
        End If
        dv.Relationship(hs) = Enums.eRelationship.Child
      End If

      '
      ' Clear the value status pairs
      '
      hs.DeviceVSP_ClearAll(dv_ref, True)
      hs.DeviceVGP_ClearAll(dv_ref, True)
      hs.SaveEventsDevices()

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      Dim VSPair As VSPair
      Dim VGPair As VGPair

      Select Case dv_type
        Case "Netatmo State"

          DevicePairs.Clear()
          DevicePairs.Add(New hspi_device_pairs(-3, "", "state_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-2, "Disable", "state_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Enable", "state_enabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Disabled", "state_disabled.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(1, "Enabled", "state_enabled.png", HomeSeerAPI.ePairStatusControl.Status))

          '
          ' Add the Status Graphic Pairs
          '
          For Each Pair As hspi_device_pairs In DevicePairs

            VSPair = New VSPair(Pair.Type)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = Pair.Value
            VSPair.Status = Pair.Status
            VSPair.Render = Enums.CAPIControlType.Values
            hs.DeviceVSP_AddPair(dv_ref, VSPair)

            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = Pair.Value
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
            hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Next

          Dim dev_status As Integer = IIf(gMonitoring = True, 1, 0)
          hs.SetDeviceValueByRef(dv_ref, dev_status, False)

          DeviceShowValues = True

      End Select

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreatePluinDevice()")
      Return "Failed to create HomeSeer device due to error."
    End Try

  End Function

  ''' <summary>
  ''' Create the HomeSeer Root Device
  ''' </summary>
  ''' <param name="dev_root_id"></param>
  ''' <param name="dev_root_name"></param>
  ''' <param name="dev_root_type"></param>
  ''' <param name="dev_root_addr"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateRootDevice(ByVal dev_root_id As String,
                                     ByVal dev_root_name As String,
                                     ByVal dev_root_type As String,
                                     ByVal dev_root_addr As String,
                                     ByVal dv_ref_child As Integer) As Integer

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Set the local variables
      '
      If dev_root_id = "Plugin" Then
        dv_name = "Netatmo Plugin"
        dv_addr = String.Format("{0}-Root", dev_root_name)
        dv_type = dev_root_type
      Else
        dv_name = dev_root_name
        dv_addr = dev_root_addr
        dv_type = dev_root_type
      End If

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} root device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} root device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = IIf(dev_root_id = "Plugin", "Plug-ins", dv_type)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this a parent root device
      '
      dv.Relationship(hs) = Enums.eRelationship.Parent_Root
      dv.AssociatedDevice_Add(hs, dv_ref_child)

      Dim image As String = "device_root.png"

      Dim VSPair As VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Root"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, image)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception

    End Try

    Return dv_ref

  End Function

  ''' <summary>
  ''' Subroutine to create the HomeSeer device
  ''' </summary>
  ''' <param name="dv_root_name"></param>
  ''' <param name="dv_root_type"></param>
  ''' <param name="dv_root_addr"></param>
  ''' <param name="dv_name"></param>
  ''' <param name="dv_type"></param>
  ''' <param name="dv_addr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetHomeSeerDevice(ByVal device_id As String,
                                    ByVal module_id As String,
                                    ByVal dv_root_name As String,
                                    ByVal dv_root_type As String,
                                    ByVal dv_root_addr As String,
                                    ByVal dv_name As String,
                                    ByVal dv_type As String,
                                    ByVal dv_addr As String) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim DevicePairs As New ArrayList
    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Define local variables
      '
      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Make this device a child of the root
      '
      If dv.Relationship(hs) <> Enums.eRelationship.Child Then

        If bCreateRootDevice = True Then
          dv.AssociatedDevice_ClearAll(hs)
          Dim dvp_ref As Integer = CreateRootDevice("", dv_root_name, dv_root_type, dv_root_addr, dv_ref)
          If dvp_ref > 0 Then
            dv.AssociatedDevice_Add(hs, dvp_ref)
          End If
          dv.Relationship(hs) = Enums.eRelationship.Child
        End If

        hs.SaveEventsDevices()
      End If

      '
      ' Exit if our device exists
      '
      If bDeviceExists = True Then Return dv_addr

      '
      ' Store the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("device_id", device_id)
      pdata.AddNamed("module_id", module_id)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = dv_root_type

        '
        ' Update the last change date
        ' 
        dv.Last_Change(hs) = DateTime.Now
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Clear the value status pairs
      '
      hs.DeviceVSP_ClearAll(dv_ref, True)
      hs.DeviceVGP_ClearAll(dv_ref, True)
      hs.SaveEventsDevices()

      Dim VSPair As VSPair
      Dim VGPair As VGPair
      Select Case dv_type
        Case "Netatmo Firmware"
          '
          ' Netatmo Firmware
          '
          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = Integer.MaxValue
          VSPair.RangeStatusPrefix = "Firmware "
          VSPair.Render = Enums.CAPIControlType.Not_Specified
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = Integer.MaxValue
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "netatmo_firmware.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Netatmo Modules"
          '
          ' Netatmo Modules Device
          '
          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 10
          VSPair.RangeStatusSuffix = " Modules"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = 10
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "netatmo_modules.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Netatmo Update"
          '
          ' Netatmo Update Device
          '
          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = Long.MaxValue
          VSPair.RangeStatusSuffix = " Minutes Ago"
          VSPair.IncludeValues = True
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = Long.MaxValue
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "time.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Netatmo Setpoint Time"
          '
          ' Netatmo Setpoint Time Devices
          '
          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "N/A"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 1
          VSPair.RangeEnd = Long.MaxValue
          VSPair.RangeStatusPrefix = "Within "
          VSPair.RangeStatusSuffix = " Minutes"
          VSPair.IncludeValues = True
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = Long.MaxValue
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "time.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Netatmo RF"
          '
          ' Add VSPair (Low)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 90
          VSPair.RangeEnd = 1000
          VSPair.RangeStatusSuffix = "Low"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Low)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 90
          VGPair.RangeEnd = 1000
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "signal_low.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (Medium)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 80
          VSPair.RangeEnd = 89
          VSPair.RangeStatusSuffix = "Medium"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Medium)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 80
          VGPair.RangeEnd = 89
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "signal_medium.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (High)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 70
          VSPair.RangeEnd = 79
          VSPair.RangeStatusSuffix = "High"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (High)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 70
          VGPair.RangeEnd = 79
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "signal_high.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (Full)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 69
          VSPair.RangeStatusSuffix = "Full"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Full)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = 69
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "signal_full.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Netatmo WiFi"
          '
          ' Add VSPair (Good)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 56
          VSPair.RangeStatusSuffix = "Good"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Good)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = 56
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "wifi_good.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (Average)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 57
          VSPair.RangeEnd = 70
          VSPair.RangeStatusSuffix = "Average"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Average)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 57
          VGPair.RangeEnd = 70
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "wifi_average.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (Bad)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 71
          VSPair.RangeEnd = 10000
          VSPair.RangeStatusSuffix = "Bad"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Bad)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 71
          VGPair.RangeEnd = 1000
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "wifi_bad.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Netatmo Battery"
          Select Case dv_name
            Case "Battery Status"
              '
              ' Add VSPair (Low)
              '
              VSPair = New VSPair(ePairStatusControl.Status)
              VSPair.PairType = VSVGPairType.Range
              VSPair.RangeStart = 0
              VSPair.RangeEnd = 3000
              VSPair.RangeStatusSuffix = "Low"
              VSPair.IncludeValues = False
              VSPair.Render = Enums.CAPIControlType.Values
              hs.DeviceVSP_AddPair(dv_ref, VSPair)

              '
              ' Add VSPair (Low)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 0
              VGPair.RangeEnd = 3000
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_25.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

              '
              ' Add VSPair (Medium)
              '
              VSPair = New VSPair(ePairStatusControl.Status)
              VSPair.PairType = VSVGPairType.Range
              VSPair.RangeStart = 3001
              VSPair.RangeEnd = 3599
              VSPair.RangeStatusSuffix = "Medium"
              VSPair.IncludeValues = False
              VSPair.Render = Enums.CAPIControlType.Values
              hs.DeviceVSP_AddPair(dv_ref, VSPair)

              '
              ' Add VSPair (Medium)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 3001
              VGPair.RangeEnd = 3599
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_50.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

              '
              ' Add VSPair (High)
              '
              VSPair = New VSPair(ePairStatusControl.Status)
              VSPair.PairType = VSVGPairType.Range
              VSPair.RangeStart = 3600
              VSPair.RangeEnd = 10000
              VSPair.RangeStatusSuffix = "High"
              VSPair.IncludeValues = False
              VSPair.Render = Enums.CAPIControlType.Values
              hs.DeviceVSP_AddPair(dv_ref, VSPair)

              '
              ' Add VSPair (High)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 3600
              VGPair.RangeEnd = 10000
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_100.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

            Case "Battery Percentage"
              '
              ' Add VSPair (Low)
              '
              VSPair = New VSPair(ePairStatusControl.Status)
              VSPair.PairType = VSVGPairType.Range
              VSPair.RangeStart = 0
              VSPair.RangeEnd = 100
              VSPair.RangeStatusSuffix = "%"
              VSPair.IncludeValues = True
              VSPair.Render = Enums.CAPIControlType.Values
              hs.DeviceVSP_AddPair(dv_ref, VSPair)

              '
              ' Add VSPair (0%)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 0
              VGPair.RangeEnd = 10
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_0.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

              '
              ' Add VSPair (25%)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 11
              VGPair.RangeEnd = 25
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_25.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

              '
              ' Add VSPair (50%)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 26
              VGPair.RangeEnd = 50
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_50.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

              '
              ' Add VSPair (75%)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 51
              VGPair.RangeEnd = 75
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_75.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

              '
              ' Add VSPair (100%)
              '
              VGPair = New VGPair()
              VGPair.PairType = VSVGPairType.Range
              VGPair.RangeStart = 76
              VGPair.RangeEnd = 100
              VGPair.Graphic = String.Format("{0}{1}", gImageDir, "battery_100.png")
              hs.DeviceVGP_AddPair(dv_ref, VGPair)

          End Select

        Case "Netatmo Temperature"
          '
          ' Format Temperature Status Suffix
          '
          Dim strStatusSuffix As String = GetDeviceSuffix("", dv_type)

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = -67
          VSPair.RangeEnd = 257
          VSPair.RangeStatusPrefix = ""
          VSPair.RangeStatusSuffix = strStatusSuffix
          VSPair.RangeStatusDecimals = 1
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = -67
          VGPair.RangeEnd = 257
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "temperature.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Netatmo Setpoint"
          '
          ' Format Setpoint Status Suffix
          '
          Dim strStatusSuffix As String = GetDeviceSuffix("", dv_type)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "N/A"
          VSPair.Render = Enums.CAPIControlType.Values
          VSPair.ControlUse = ePairControlUse.Not_Specified
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "temperature.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          For value As Double = 7 To 30 Step 0.5
            VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = value
            VSPair.Status = String.Format("{0} {1}", value.ToString, strStatusSuffix)
            VSPair.Render = Enums.CAPIControlType.Values
            VSPair.ControlUse = ePairControlUse._HeatSetPoint
            hs.DeviceVSP_AddPair(dv_ref, VSPair)
          Next

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 7
          VGPair.RangeEnd = 30
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "temperature.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          DeviceShowValues = True

        Case "Netatmo Setpoint Mode"
          '
          ' Rachio ScheduleRule Status
          '
          DevicePairs.Clear()
          DevicePairs.Add(New hspi_device_pairs(0, "Off", "mode_off.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(1, "Program", "mode_program.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(2, "Away", "mode_away.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(3, "Frost-guard", "mode_hg.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(4, "Manual", "mode_manual.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(5, "Max", "mode_max.png", HomeSeerAPI.ePairStatusControl.Both))

          '
          ' Add the Status Graphic Pairs
          '
          For Each Pair As hspi_device_pairs In DevicePairs

            VSPair = New VSPair(Pair.Type)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = Pair.Value
            VSPair.Status = Pair.Status
            VSPair.Render = Enums.CAPIControlType.Values
            hs.DeviceVSP_AddPair(dv_ref, VSPair)

            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = Pair.Value
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
            hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Next

          DeviceShowValues = True

      End Select

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "GetHomeSeerDevice()")
    End Try

    Return dv_addr

  End Function

  ''' <summary>
  ''' Returns the HomeSeer Device Address
  ''' </summary>
  ''' <param name="Address"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDeviceAddress(ByVal Address As String) As String

    Try

      Dim dev_ref As Integer = hs.DeviceExistsAddress(Address, False)
      If dev_ref > 0 Then
        Return Address
      Else
        Return ""
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "GetDevCode")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Locates device by device code
  ''' </summary>
  ''' <param name="srDeviceCode"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByCode(ByVal srDeviceCode As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.GetDeviceRef(srDeviceCode)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "LocateDeviceByCode")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_value"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal dv_addr As String,
                              ByVal dv_value As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_value), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        If IsNumeric(dv_value) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblNewValue As Double = Double.Parse(dv_value)

          If dblDeviceValue <> dblNewValue Then
            hs.SetDeviceValueByRef(dv_ref, dblNewValue, True)
          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Sets the HomeSeer device string
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_string"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceString(ByVal dv_addr As String,
                               ByVal dv_string As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_string), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        hs.SetDeviceString(dv_ref, dv_string, True)

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceString()")

    End Try

  End Sub

  ''' <summary>
  ''' Returns the Device Units
  ''' </summary>
  ''' <param name="deviceKey"></param>
  ''' <param name="deviceType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDeviceSuffix(deviceKey As String, deviceType As String) As String

    Dim strStatusSuffix As String = String.Empty

    Try

      Select Case deviceType
        Case "Temperature Trend", "Pressure Trend"
        Case "Netatmo Battery"
        Case "Netatmo Temperature", "Netatmo Setpoint"
          strStatusSuffix = String.Concat(" °", IIf(gTempUnit = 0, "C", "F"))

        Case "Netatmo CO2"
          strStatusSuffix = " ppm"

        Case "Netatmo Humidity"
          strStatusSuffix = String.Concat(" ", "%")

        Case "Netatmo Noise"
          strStatusSuffix = " dB"

        Case "Netatmo Rain"
          strStatusSuffix = String.Concat(" ", IIf(gRainUnit = 0, "mm", """"))

        Case "Netatmo Wind"
          strStatusSuffix = String.Concat(" ", IIf(gWindUnit = 0, "kph", "mph"))

          Select Case gWindUnit
            Case 0 : strStatusSuffix = String.Concat(" ", "kph")
            Case 1 : strStatusSuffix = String.Concat(" ", "mph")
            Case 2 : strStatusSuffix = String.Concat(" ", "ms")
            Case 3 : strStatusSuffix = String.Concat(" ", "kph")    ' Beaufort 
            Case 4 : strStatusSuffix = String.Concat(" ", "knot")
          End Select

        Case "Netatmo Pressure"
          Select Case gPressureUnit
            Case 0 : strStatusSuffix = String.Concat(" ", "mb")
            Case 1 : strStatusSuffix = String.Concat(" ", "inHg")
            Case 2 : strStatusSuffix = String.Concat(" ", "mmHg")
          End Select

      End Select

    Catch pEx As Exception

    End Try

    Return strStatusSuffix

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="strKeyName"></param>
  ''' <param name="strKeyValue"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal strStationNumber As String,
                              ByVal strKeyName As String,
                              ByVal strKeyValue As String)

    Try
      '
      ' Check to see if we need to adjust the device value
      '
      If String.IsNullOrEmpty(strKeyValue) = True Then strKeyValue = "--"
      If strKeyValue.Length = 0 Then strKeyValue = "--"

      Dim dv_type As String = ""  'GetWeatherType(strKeyName)
      Dim dv_addr As String = ""  'Stations(strStationNumber)(strKeyName)("DevCode")
      Dim strKeyUnits As String = GetDeviceSuffix(strKeyName, dv_type)

      Dim strDeviceImage As String = FormatDeviceImage(strKeyName, strKeyValue, gStatusImageSizeHeight, gStatusImageSizeWidth)
      Dim strDeviceIcon As String = FormatDeviceImage(strKeyName, strKeyValue, 16, 16)

      Dim strDeviceValue As String = FormatDeviceValue(strKeyName, strKeyValue, strKeyUnits)
      Dim strDeviceString As String = FormatDeviceString(strDeviceImage, strDeviceValue)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      If bDeviceExists = True Then
        '
        ' Update the HomeSeer device value
        '
        Dim dblDeviceValue As Double = 0
        If Double.TryParse(strKeyValue, System.Globalization.NumberStyles.Float, nfi, dblDeviceValue) = True Then

          If hs.DeviceString(dv_ref) <> "" Then
            hs.SetDeviceString(dv_ref, "", False)
          End If

          If hs.DeviceValue(dv_ref) <> dblDeviceValue Then
            SetDeviceValue(dv_addr, dblDeviceValue)
          End If

        Else
          '
          ' Update the Device String
          '
          If hs.DeviceString(dv_ref) <> strDeviceString Then
            hs.SetDeviceString(dv_ref, strDeviceString, True)
          End If

        End If

      End If

      '
      ' Update the stations hashtable
      '
      'If Stations(strStationNumber)(strKeyName)("Value") <> strKeyValue Then

      '  Stations(strStationNumber)(strKeyName)("Image") = strDeviceImage
      '  Stations(strStationNumber)(strKeyName)("Icon") = strDeviceIcon
      '  Stations(strStationNumber)(strKeyName)("Units") = strKeyUnits
      '  Stations(strStationNumber)(strKeyName)("Value") = strKeyValue
      '  Stations(strStationNumber)(strKeyName)("String") = strDeviceValue

      '  Stations(strStationNumber)(strKeyName)("LastChange") = Now.ToString
      'End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Formats the HomeSeer Device Image
  ''' </summary>
  ''' <param name="strDeviceKey"></param>
  ''' <param name="strDeviceValue"></param>
  ''' <param name="imageHeight"></param>
  ''' <param name="imageWidth"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FormatDeviceImage(ByVal strDeviceKey As String,
                                      ByVal strDeviceValue As String,
                                      Optional ByVal imageHeight As Integer = 16,
                                      Optional ByVal imageWidth As Integer = 16) As String

    Dim strDeviceImage As String = "unknown.png"

    If gDeviceImage = False Then Return strDeviceImage

    Try

      ' "Weather", "Temperature", "Humidity", "Wind", "Rain", "Pressure", "Visibility", "Forecast", "Alerts"

      Dim weatherType As String = ""  'GetWeatherType(strDeviceKey)
      Select Case weatherType
        Case "Weather"
          Select Case strDeviceKey
            Case "ob-Date"
              strDeviceImage = "time.png"
            Case "current-condition"
              strDeviceImage = String.Format("cond{0}.png", strDeviceValue.PadLeft(3, "0"))
          End Select
        Case "Temperature", "Setpoint"
          strDeviceImage = "temperature.png"
      End Select

      If strDeviceImage.Length > 0 Then
        strDeviceImage = String.Format("<img height=""{0}"" align=""absmiddle"" width=""{1}"" src=""{2}{3}""/>", imageHeight.ToString, imageWidth.ToString, gImageDir, strDeviceImage)
      End If

    Catch pEx As Exception
      strDeviceImage = String.Format("<img height=""{0}"" align=""absmiddle"" width=""{1}"" src=""{2}{3}""/>", imageHeight.ToString, imageWidth.ToString, gImageDir, "unknown.png")
    End Try

    Return strDeviceImage

  End Function

  ''' <summary>
  ''' Formats the HomeSeer Device Value
  ''' </summary>
  ''' <param name="strDeviceKey"></param>
  ''' <param name="strDeviceValue"></param>
  ''' <param name="strDeviceUnits"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FormatDeviceValue(ByVal strDeviceKey As String,
                                      ByVal strDeviceValue As String,
                                      ByVal strDeviceUnits As String) As String


    Dim strDeviceString As String = ""

    Try

      If strDeviceUnits.Length > 0 Then
        strDeviceString = String.Format("{0} {1}", strDeviceValue, strDeviceUnits)
      Else
        strDeviceString = strDeviceValue
      End If

    Catch ex As Exception

    End Try

    Return strDeviceString

  End Function

  ''' <summary>
  ''' Formats the HomeSeer Device String
  ''' </summary>
  ''' <param name="strDeviceImage"></param>
  ''' <param name="strDeviceValue"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FormatDeviceString(ByVal strDeviceImage As String,
                                       ByVal strDeviceValue As String) As String


    Dim strDeviceString As String = ""

    Try

      If strDeviceImage.Length > 0 Then
        strDeviceString = String.Format("{0}&nbsp;{1}", strDeviceImage, strDeviceValue)
      Else
        strDeviceString = strDeviceValue
      End If

    Catch pEx As Exception

    End Try

    Return strDeviceString

  End Function

End Module
