Imports System.Diagnostics
Imports System.Reflection
Imports System.Resources
Imports System.Runtime.InteropServices

Friend Module VisualBasic6_Support

#Region "Initialization"

   ' Constants for the App6 object
   Public Const APP6_HELPFILE As String = ""
   Public Const APP6_HELPCONTEXTID As Integer = 0
   Public Const APP6_TITLE As String = "calm_fire_toolbar"
   Private Const APP6_UNIQUEID As Long = 21290932764571444           ' <== don't modify!

	' Initialize the VB6 support library
	Sub New()
		' Uncomment next statement to skip assembly exploration at startup
		' CodeArchitects.VB6Library.VB6Config.ParseAssembliesAtStartup = False 
		CodeArchitects.VB6Library.InitializeLibrary(Application.OpenForms, My.Application.Info, APP6_HELPFILE, APP6_HELPCONTEXTID, APP6_UNIQUEID)
		' next statement can be commented out in applications that don't use DDE
		CodeArchitects.VB6Library.VB6Config.DDEAppTitle = APP6_TITLE

	End Sub

   Public Sub EnsureVB6LibraryInitialization()
		' The only purpose of this empty method is forcing the execution of the static 
		' constructor the first time the method is invoked
	End Sub

#End Region

#Region "Default instances"

	' Default instances of global classes used by this project
	Private m_Outlook_Application_DefInstance As Outlook.Application

	Public ReadOnly Property Outlook_Application_DefInstance() As Outlook.Application
		Get
			If m_Outlook_Application_DefInstance Is Nothing Then m_Outlook_Application_DefInstance = New Outlook.Application()
			Return m_Outlook_Application_DefInstance
		End Get
	End Property

#End Region

#Region "Resources support"

   ' support for reading resources

   Dim resourceMan As ResourceManager = New Global.System.Resources.ResourceManager(GetType(VisualBasic6_Support).Namespace & ".Resources", GetType(VisualBasic6_Support).Assembly)

   Public Function LoadResString6(ByVal id As Object) As String
      Return DirectCast(GetResourceFromID("str", id), String)
   End Function

   Public Function LoadResPicture6(ByVal id As Object, ByVal resType As Integer) As Object
      Dim prefix As String = Microsoft.VisualBasic.Choose(resType + 1, "bmp", "ico", "cur")
      Return GetResourceFromID(prefix, id)
   End Function

   Public Function LoadResData6(ByVal id As Object, ByVal resType As Object) As Object
      Dim prefix As String = "?"
      If TypeOf resType Is String Then
         prefix = resType.ToString() & "_"
      ElseIf IsNumeric(resType.ToString()) Then
         prefix = Microsoft.VisualBasic.Choose(CInt(resType), "cur", "bmp", "ico", "?", "?", "str")
      End If
      Return GetResourceFromID(prefix, id)
   End Function

   ' extract a resource, throw if not found

   Private Function GetResourceFromID(ByVal prefix As String, ByVal id As Object) As Object
      ' ensure ID doesn't contain invalid chars
      Dim strId As String = System.Text.RegularExpressions.Regex.Replace(id.ToString(), "[^A-Za-z0-9_.]", "_")
      Dim res As Object = resourceMan.GetObject(prefix & strId, Nothing)

      If res IsNot Nothing Then
         Select Case prefix
            Case "str" : If Not TypeOf res Is String Then res = Nothing
            Case "bmp" : If Not TypeOf res Is System.Drawing.Image Then res = Nothing
            Case "ico" : If Not TypeOf res Is System.Drawing.Image AndAlso Not TypeOf res Is Icon Then res = Nothing
            Case "cur" : If Not TypeOf res Is Cursor Then res = Nothing
         End Select
      Else
         ' Try again with all-uppercase ID
         res = resourceMan.GetObject(prefix.ToUpper() & strId, Nothing)
      End If
      ' Check whether the resource is missing or is of the wrong type
      If res Is Nothing Then Err.Raise(326, , "Resource with identifier '" & strId & "' not found")
      Return res
   End Function

#End Region

#Region "Debug support"

   <Conditional("DEBUG")> _
   Public Sub DebugPrintLine6(ByVal ParamArray args() As Object)
      CodeArchitects.VB6Library.DebugPrintLine6(args)
   End Sub

   <Conditional("DEBUG")> _
   Public Sub DebugPrint6(ByVal ParamArray args() As Object)
      CodeArchitects.VB6Library.DebugPrint6(args)
   End Sub

#End Region

End Module

