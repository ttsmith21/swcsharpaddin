Option Explicit On
Friend Module MCommon
    Public Const MaxDouble As Double = 1.79769313486231E+308
    Public Const MinDouble As Double = -1.79769313486231E+308

    Public Const strProductName As String = "Extract Profile Data"
    Public strProductVersion As String = ""
    Public Const strRoundShape As String = "Round"
    Public Const strRectangleShape As String = "Rectangle"
    Public Const strSquareShape As String = "Square"
    Public Const strAngleShape As String = "Angle"
    Public Const strChannelShape As String = "Channel"
    Public Const strNoShape As String = ""

    Public Const strOutputfolderName As String = ""
    Public Enum EnumShape
        none = 0
        round = 1
        square = 2
        rectangle = 3
        angle = 4
        channel = 5
    End Enum
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As IntPtr) As IntPtr
    Public Function IsTendsToZero(ByRef val As Double) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Math.Abs(val) < 0.000000001 Then
            bReturn = True
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function GetParentFolderName(ByRef strFilePathName As String) As String
        On Error Resume Next
        Dim strReturn As String = ""
        If Not String.IsNullOrEmpty(strFilePathName) Then
            If IO.File.Exists(strFilePathName) Then
                strReturn = FileIO.FileSystem.GetFileInfo(strFilePathName).DirectoryName
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return strReturn
    End Function
    Public Function GetFileNameWithoutExtensionFromFilePathName(ByRef strFilePathName As String) As String
        On Error Resume Next
        Dim strReturn As String = ""
        If FileIO.FileSystem.FileExists(strFilePathName) Then
            strReturn = IO.Path.GetFileNameWithoutExtension(strFilePathName)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return strReturn
    End Function
    Public Function GetProductVersion() As String
        On Error Resume Next
        If String.IsNullOrEmpty(strProductVersion) Then
            strProductVersion = String.Format("Version {0}", My.Application.Info.Version.ToString)
        End If
        Return strProductVersion
    End Function
End Module
