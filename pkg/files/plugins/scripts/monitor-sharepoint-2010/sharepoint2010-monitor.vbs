Set objWSH =  CreateObject("WScript.Shell")
Set WshSysEnv = objWSH.Environment("Process") 
strComputer = WshSysEnv("UPTIME_HOSTNAME")
strUser = WshSysEnv("UPTIME_USERNAME")
strPassword = WshSysEnv("UPTIME_PASSWORD")
On Error Resume Next


Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, "\root\CIMV2", strUser, strPassword, "ms_409")
objSWbemServices.Security_.ImpersonationLevel = 3

if objSWbemServices Is Nothing Then
	WScript.Echo "Access Denied! - Check your username & password"
	WScript.Quit 1
End If

Set colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MicrosoftWindowsSharePointMicrosoftSharePointFoundation4_SandboxedCodeProcessPool", "WQL", wbemFlagReturnImmediately)

if colItems is Nothing or colItems.Count = 0 Then
	WScript.Echo "Unable to SharePoint Performance Metrics via WMI"
	WScript.Quit 1
End If

IF colItems.Count <> 0 THEN
	For Each objItem In colItems
		 WScript.Echo "RequestsActive " & objItem.RequestsActive
		 WScript.Echo "RequestsQueued " & objItem.RequestsQueued
		 WScript.Echo "RequestsTimedoutTotal " & objItem.RequestsTimedoutTotal
	Next
END IF

Set colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MicrosoftWindowsSharePointMicrosoftSharePointFoundation4_SharePointFoundationSearchGatherer", "WQL", wbemFlagReturnImmediately)

if colItems is Nothing or colItems.Count = 0 Then
	WScript.Echo "Unable to SharePoint Performance Metrics via WMI"
	WScript.Quit 1
End If

IF colItems.Count <> 0 THEN
	For Each objItem In colItems
		WScript.Echo "ActiveQueueLength " & objItem.ActiveQueueLength
		WScript.Echo "IdleThreads " & objItem.IdleThreads
		WScript.Echo "ThreadsAccessingNetwork " & objItem.ThreadsAccessingNetwork
		WScript.Echo "FilteringThreads " & objItem.FilteringThreads
		WScript.Echo "ThreadsInPlugins " & objItem.ThreadsInPlugins		
	NEXT
END IF 

Set colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MicrosoftWindowsSharePointMicrosoftSharePointFoundation4_SharePointFoundationSearchGathererProjects", "WQL", wbemFlagReturnImmediately)

if colItems is Nothing or colItems.Count = 0 Then
	WScript.Echo "Unable to SharePoint Performance Metrics via WMI"
	WScript.Quit 1
End If

IF colItems.Count <> 0 THEN
	For Each objItem In colItems
      		WScript.Echo "Crawlsinprogress " & objItem.Crawlsinprogress
	NEXT
END IF
