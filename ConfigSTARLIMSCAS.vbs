Call RunAndReport("192.168.1.187", "2.0.50727")

Sub RunAndReport(site, strVersions)
	If WScript.Arguments.Count = 0 Then 
		Call RunAndReport2(site, strVersions)
	Else  
		Call RunAndReport2(site, strVersions) 
	End If
    
    EnableIEHosting()           
End Sub

Sub RunAndReport2(site, strVersions)
	
	Dim strPass, strFail	
	Dim strFinalResult, strRunResult
	Dim intIcon
	Dim arVersions
	Dim strVer

	arVersions = Split(strVersions, ",")

	strFinalResult = "" 
	strPass = ""
	strFail = ""  
	intIcon = 0

	For Each strVer In arVersions
		strRunResult = RunCasPol(site, strVer)
		
		If strRunResult = "" Then
			strPass = strPass + "> v" + strVer + Chr(13) 
		Else
			strFail = strFail + "> v" + strVer + ": " + strRunResult + Chr(13)  
		End If
	Next

	If Len(strPass) > 0 Then
		strFinalResult = strFinalResult + _  
			"Browser configured succesfully for running STARLIMS software on the following .NET Frameworks." + _
		 	Chr(13) + "New STARLIMS codegroups were added to your .NET machine settings." + Chr(13) + Chr(13) + _
			strPass + Chr(13) + Chr(13)
		
		intIcon = 64
	End If
    
	If Len(strFail) > 0 Then
		strFinalResult = strFinalResult + _ 
			"Browser configuration failed on the following .NET frameworks." + Chr(13) + Chr(13) + _
			strFail + Chr(13) + Chr(13)
		
		If intIcon = 0 Then intIcon = 16 Else intIcon = 48  
	End If

End Sub

Function RunCasPol(site, strVer)
	
	Dim objShell  
	Dim strCasPolArgs   
	Dim strCasPolExe    
	Dim strCommandLine
	Dim intRunStatus

	strCasPolArgs = BuildCommandArguments(site)   
	
	Set objShell = WScript.CreateObject("WScript.Shell")  
	strCasPolExe = objShell.ExpandEnvironmentStrings("%windir%\Microsoft.NET\Framework\v" + strVer + "\caspol.exe")
	
	strCommandLine = strCasPolExe + " -polchgprompt off"
	Call objShell.Run(strCommandLine, 0, True)   

	strCommandLine = strCasPolExe + " -rg ""STARLIMS " + site + """" 
	Call DelExistingSTARLIMSCodeGroups(objShell, strCommandLine)

	strCommandLine = strCasPolExe + " " + strCasPolArgs   
	'intRunStatus = objShell.Run(strCommandLine, 0, True)  
	
	Set objShell = Nothing
	
	RunCasPol = ""  
End Function

Function BuildCommandArguments(site)
	
	BuildCommandArguments = "-m -ag 1 -site " + site + " FullTrust -n ""STARLIMS " + site + """ -d ""This code group grants the FullTrust permission set to assemblies of STARLIMS software."""  
End Function

Sub DelExistingSTARLIMSCodeGroups(objShell, strCommandLine)
	
	'intRunStatus = 0
	'Do While intRunStatus = 0
	' intRunStatus = objShell.Run(strCommandLine, 0, True)
	'Loop
		
End Sub

Sub EnableIEHosting()
    
    Set Shell = CreateObject( "WScript.Shell" )  
	
	Shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework\EnableIEHosting", 1, "REG_DWORD"   
	Shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\EnableIEHosting", 1, "REG_DWORD"
End Sub