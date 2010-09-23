' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Module modExtension

    Public Const IncidentMessage As String = _
    "Incident is not open. Ensure that 1) the DEC Fire Extension is turned on (Tools/Extensions) and 2) " & _
    "that you have logged on to the incident (logon button on far left of fire incident toolbar) and 3) " & _
    "that you have either created or opened an incident (Incident menu pulldown, 2nd left on fire incident" & _
    "toolbar).  Only after all these things have been done can you fight fires with maps!"

    Public Const LoginMessage As String = "Please turn on the DEC Fire Extension (Tools/Extensions)"


	' UPGRADE_INFO (#0561): The 'DecConfigFile' symbol was defined without an explicit "As" clause.
	Private Const DecConfigFile As String = "V:\GIS1-Corporate\Resource\AppConfigs\FireConfig.txt"
	'Const DecConfigFile = "C:\Temp\FireConfig.txt"

	' UPGRADE_INFO (#0551): The 'Key' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function GetConfigValue(ByRef Key As String) As String
		Dim FSO As New VB6FileSystemObject
		Dim ts As VB6TextStream = FSO.OpenTextFile(DecConfigFile)
		Dim strOutput As String = ""
		Dim strLine As String = ""
		Dim beginPos As Short
		Dim endPos As Short
		Dim stringpos As Short

		Do Until ts.AtEndOfStream
			
			strLine = ts.ReadLine()
			
			stringpos = InStr(1, strLine, Key, CompareMethod.Text)
			
			If stringpos <> 0 Then
				stringpos = InStr(stringpos, strLine, "=", CompareMethod.Text)
				If stringpos <> 0 Then
					beginPos = stringpos + 1
					endPos = Len6(strLine)
					strOutput = Mid(strLine, beginPos, endPos)
					Exit Do
				End If
			End If
		Loop
		ts.Close()
		Return strOutput
		
	End Function

End Module
