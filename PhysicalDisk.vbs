'###################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								  		      ##
'##																													  		      ##
'## December, 2014																									  		      ##
'##																													  		      ##
'## Version 1.0																										  		      ##
'##																													  		      ##
'## DESCRIPTION: Monitor network interface traffic and errors																	  ##
'##																													  		      ##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "PhysicalDisk.vbs" <HOST> <METRIC_STATE> <USERNAME> <PASSWORD> <DOMAIN>    ##
'##																													  		      ##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "PhysicalDisk.vbs" "10.10.10.1" "1,1,1,0,1,0,1,1,1" "user" "pwd" "domain" ##
'##																													              ##
'## README:	<METRIC_STATE> is generated internally by Tellki and its only used by Tellki default monitors. 						  ##
'##         1 - metric is on ; 0 - metric is off					              												  ##
'## 																												              ##
'## 	    <USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this ##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   ##
'##			pass them to the script.																						      ##
'## 																												              ##
'###################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 5 Then 
	CALL ShowError(3, 0)
End If
'Set Culture - en-us
SetLocale(1033)

'METRIC_ID
Const DiskTimePerc = "164:% Disk Time:6"
Const DiskTransSec = "152:Disk Transfers/Sec:4"
Const AvgDiskBytesTrans = "81:Average Disk Bytes/Transfer:4"
Const DiskBytesSec = "191:Disk Bytes/Sec:4"
Const AvgDiskSecTrans = "141:Average Disk sec/Transfer:4"
Const DiskWriteTransSec = "225:Disk Writes Transfers/Sec:4"
Const DiskReadTransSec = "226:Disk Reads Transfers/Sec:4"
Const DiskWriteBytesSec = "227:Disk Writes Bytes/Sec:4"
Const DiskReadBytesSec = "228:Disk Reads Bytes/Sec:4"

'INPUTS
Dim Host, MetricState, Username, Password, Domain
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
Username = WScript.Arguments(2)
Password = WScript.Arguments(3)
Domain = WScript.Arguments(4)

Dim arrMetrics
arrMetrics = Split(MetricState,",")
Dim objSWbemLocator, objSWbemServices, colItems
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim Counter, objItem, FullUserName
Counter = 0

	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
		if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)
	
	if Err.Number = 0 Then
		if IsObject(objSWbemServices) = True Then
			objSWbemServices.Security_.ImpersonationLevel = 3
			Dim OS
			OS = GetOSVersion(objSWbemServices)
			if OS >= 4000 Then
				Set colItems = objSWbemServices.ExecQuery( _
					"SELECT Name,PercentDiskTime,DiskTransfersPersec,AvgDiskBytesPerTransfer,DiskBytesPersec,AvgDisksecPerTransfer,DiskWritesPerSec,DiskReadsPerSec,DiskWriteBytesPerSec,DiskReadBytesPerSec from Win32_PerfFormattedData_PerfDisk_PhysicalDisk WHERE Name <> '_Total'",,16) 
				If colItems.Count <> 0 Then 	
					For Each objItem in colItems
						'PercentDiskTime
						If arrMetrics(0)=1 Then _
						CALL Output(DiskTimePerc,FormatNumber(objItem.PercentDiskTime),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'DiskTransfersPersec
						If arrMetrics(1)=1 Then _
						CALL Output(DiskTransSec,FormatNumber(objItem.DiskTransfersPersec),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'AvgDiskBytesPerTransfer
						If arrMetrics(2)=1 Then _
						CALL Output(AvgDiskBytesTrans,FormatNumber((objItem.AvgDiskBytesPerTransfer/1024)/1024),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'DiskBytesPersec
						If arrMetrics(3)=1 Then _
						CALL Output(DiskBytesSec,FormatNumber((objItem.DiskBytesPersec/1024)/1024),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'AvgDisksecPerTransfer
						If arrMetrics(4)=1 Then _
						CALL Output(AvgDiskSecTrans,FormatNumber(objItem.AvgDisksecPerTransfer),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'DiskWritesPerSec
						If arrMetrics(5)=1 Then _
						CALL Output(DiskWriteTransSec,FormatNumber(objItem.DiskWritesPerSec),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'DiskReadsPerSec
						If arrMetrics(6)=1 Then _
						CALL Output(DiskReadTransSec,FormatNumber(objItem.DiskReadsPerSec),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'DiskWriteBytesPerSec
						If arrMetrics(7)=1 Then _
						CALL Output(DiskWriteBytesSec,FormatNumber((objItem.DiskWriteBytesPerSec/1024)/1024),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
						'DiskReadBytesPerSec
						If arrMetrics(8)=1 Then _
						CALL Output(DiskReadBytesSec,FormatNumber((objItem.DiskReadBytesPerSec/1024)/1024),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
					Next
				Else
					'If there is no response in WMI query
					CALL ShowError(5, Host)
				End If
			Else
				Set colItems = objSWbemServices.ExecQuery("Select Name from Win32_PerfRawData_PerfDisk_PhysicalDisk where Name <> '_Total'",,16)
				If colItems.Count <> 0 Then
				For Each objItem in colItems
					If arrMetrics(0)=1 Then
						CALL Output(DiskTimePerc,FormatNumber(getLogicalDiskDrive(objSWbemServices ,objItem.Name)),Mid(objItem.Name,InStr(objItem.Name," ")+1,Len(objItem.Name)))
					End If 
				Next
				Else
					'If there is no response in WMI query
					CALL ShowError(5, Host)
				End If
			End If
		End If
		If Err.number <> 0 Then
			CALL ShowError(5, Host)
			Err.Clear
		End If
	End If


If Err Then 
	CALL ShowError(1,0)
Else
	WScript.Quit(0)
End If

Function getLogicalDiskDrive(SWbem, drive)
	Dim sumdisktimePCT, i, colItems, objInstance1, objInstance2, N1, D1, N2, D2, disktime 
	sumdisktimePCT=0
	For i = 1 to 5
		Set colItems = SWbem.ExecQuery("Select Name,PercentDiskTime,Timestamp_Sys100NS from Win32_PerfRawData_PerfDisk_PhysicalDisk where Name='"&drive&"'",,16)
		For Each objInstance1 in colItems
			N1 = objInstance1.PercentDiskTime
			D1 = objInstance1.TimeStamp_Sys100NS
		Next
		WScript.Sleep(1000)
		Set colItems = SWbem.ExecQuery("Select Name,PercentDiskTime,Timestamp_Sys100NS from Win32_PerfRawData_PerfDisk_PhysicalDisk where Name='"&drive&"'",,16)
		For Each objInstance2 in colItems
			N2 = objInstance2.PercentDiskTime
			D2 = objInstance2.TimeStamp_Sys100NS
		Next
		disktime = (((N2 - N1) / (D2 - D1))) * 100
		sumdisktimePCT=disktime+sumdisktimePCT
	Next	
	getLogicalDiskDrive=Round((sumdisktimePCT/10),2)
End Function

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Function GetOSVersion(SWbem)
	Dim colItems, objItem
	Set colItems = SWbem.ExecQuery("select BuildVersion from Win32_WMISetting",,16)
	For Each objItem in colItems
		GetOSVersion = CInt(objItem.BuildVersion)
	Next
End Function

Sub Output(MetricID, MetricValue, MetricObject)
	If MetricObject <> "" Then
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "|" & MetricObject & "|" 
		Else
			CALL ShowError(5, Host) 
		End If
	Else
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "|" 
		Else
			CALL ShowError(5, Host)
		End If
	End If
End Sub


