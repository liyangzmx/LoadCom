' Title: putty launcher
' Author: Yang.Li
' Data: 2018-01-02
' Version: --
' Description: 自动判断CP210x的COM端口号并打开串口

baudRate = "115200"
usbDecriptionPrefix = "Silicon Labs CP210x USB to UART Bridge"
strComputer = "."
progName = "putty.exe"

Dim j1, j2, j3
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colDevices = objWMIService.ExecQuery _
	("Select * From Win32_USBControllerDevice")
For Each objDevice in colDevices
	strDeviceName = objDevice.Dependent 
	strQuotes = Chr(34) 
	strDeviceName = Replace(strDeviceName, strQuotes, "") 
	arrDeviceNames = Split(strDeviceName, "=") 
	strDeviceName = arrDeviceNames(1) 
	Set colUSBDevices = objWMIService.ExecQuery _ 
		("Select * From Win32_PnPEntity Where DeviceID = '" & strDeviceName & "'") 
	For Each objUSBDevice in colUSBDevices 
		j1 = InStr(1, objUSBDevice.Description, usbDecriptionPrefix, vbTextCompare)
		If j1 > 0 Then
			j2 = InStr(1, objUSBDevice.PnPDeviceID, "\", vbTextCompare)
			j3 = InStr(j2 + 1, objUSBDevice.PnPDeviceID, "\", vbTextCompare)
			id = Mid(objUSBDevice.PnPDeviceID, (j2 + 1), (j3 - j2 - 1))
			subid = Mid(objUSBDevice.PnPDeviceID, (j3 + 1))
			reg_key = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\USB\" & id & "\" & subid & "\Device Parameters\PortName"
			reg_val = CreateObject("Wscript.Shell").RegRead(reg_key)
			Set fso = CreateObject("Scripting.FileSystemObject")
			tt = fso.FileExists(progName)
			If tt=true Then
				Set wshell = Wscript.CreateObject("Wscript.Shell")
				Wshell.Run "putty.exe -serial " & reg_val & " -sercfg " & baudRate & ",8,n,1,N"
			Else
				Wscript.Echo "Uart: " & objUSBDevice.Description & " is: " & reg_val
			End If
		End If
	Next
Next

