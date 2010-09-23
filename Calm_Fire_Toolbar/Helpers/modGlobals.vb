' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Public Enum g_Fire_Dept
	DEC
	FESA
End Enum

Friend Module modGlobals

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: Nathan Eaton, CALM
	'Description: This module stores global variables utilised throughout the project
	'Called By: Many
	'Calls
	'Accepts
	'Returns
	'******************************************************************************************************************
	
	Public g_Username As String = ""
	
    Public g_incidentStatus As Boolean = False

    Public Const g_TABLE_INCIDENT_INFO As String = "Fire_IncidentInfo"
	Public Const g_FIELD_PROPVALUE As String = "Prop_Value"
	Public Const g_FIELD_PROPNAME As String = "Property"
	
	Public Const g_FIRE_BOUNDARY As String = "Fire_Boundary"
	Public Const g_FIRE_ASSIGNMENT As String = "Fire_Assignment"
	Public Const g_FIRE_LINE As String = "Fire_Line"
	Public Const g_FIRE_POINT As String = "Fire_Point"
	Public Const g_FIRE_ANNO As String = "Fire_Anno"
	
	Public Const g_FIRE_BOUNDARY_NAME As String = "Fire Boundary"
	Public Const g_FIRE_ASSIGNMENT_NAME As String = "Fire Assignment"
	Public Const g_FIRE_LINE_NAME As String = "Fire Line"
	Public Const g_FIRE_POINT_NAME As String = "Fire Point"
	Public Const g_FIRE_ANNO_NAME As String = "Fire Annotation"
	
	Public Const g_FIRE_EXTENSION As String = "DEC Fire Extension"
	
	Public g_UserDept As g_Fire_Dept
	
	Public g_INCIDENT_LAYER_NAME As String = ""
	
	Public Const g_FEATURE_FIELD As String = "Feature"
	Public Const g_REGION_FIELD As String = "Region"
	Public Const g_REGIONEX_FIELD As String = "Region_ext"
	Public Const g_FIRE_NUM_FIELD As String = "Fire_Number"
	Public Const g_USERNAME_FIELD As String = "Username"
	Public Const g_USER_DEPT_FIELD As String = "User_Dept"
	Public Const g_CAPT_DATE_FIELD As String = "Capt_Date"
	Public Const g_CAPT_TIME_FIELD As String = "Capt_Time"
	Public Const g_CAPT_METHOD_FIELD As String = "Capt_Method"
	Public Const g_HECTARE_FIELD As String = "AutoHa"
	Public Const g_PERIM_FIELD As String = "AutoPerim"
	

End Module
