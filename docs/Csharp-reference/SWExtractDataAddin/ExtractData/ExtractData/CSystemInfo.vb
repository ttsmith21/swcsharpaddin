Option Explicit On
Imports System.Management
Friend Class CSystemInfo
    Private strProcessorID As String = ""
    Private strBIOSVersion As String = ""
    Private strMACAddress As String = ""
    Private strMotherBoardID As String = ""
    Private strBIOSSerialNumber As String = ""

    Private Shared myInstance As CSystemInfo
    Public Shared Function GetInstance() As CSystemInfo
        If myInstance Is Nothing Then
            myInstance = New CSystemInfo
        End If
        Return myInstance
    End Function
    Friend Function GetProcessorId() As String
        ''Dim strProcessorId As String = String.Empty
        If String.IsNullOrEmpty(Trim(strProcessorID)) Then
            Dim query As Management.SelectQuery
            Dim search As Management.ManagementObjectSearcher
            Dim info As Management.ManagementObject
            Try
                query = New Management.SelectQuery("Win32_processor")
                Try
                    search = New Management.ManagementObjectSearcher(query)
                    For Each info In search.Get()
                        strProcessorID = info("processorId").ToString()
                    Next
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            query = Nothing
            search = Nothing
            info = Nothing
        End If

        If Err.Number <> 0 Then Err.Clear()
        Return strProcessorID
    End Function
    Friend Function GetMACAddress() As String
        If String.IsNullOrEmpty(Trim(strMACAddress)) Then
            Dim mc As ManagementClass
            Dim moc As ManagementObjectCollection
            Try
                mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
                Try
                    moc = mc.GetInstances()
                    For Each mo As ManagementObject In moc
                        If CBool(mo("IPEnabled")) Then strMACAddress = mo("MacAddress").ToString()
                        mo.Dispose()
                        strMACAddress = strMACAddress.Replace(":", String.Empty)
                    Next
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            mc = Nothing
            moc = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return strMACAddress
    End Function
    'Friend Function GetVolumeSerial(Optional ByVal strDriveLetter As String = "C") As String
    '    Dim disk As ManagementObject = New ManagementObject(String.Format("win32_logicaldisk.deviceid=""{0}:""", strDriveLetter))
    '    disk.Get()
    '    Return disk("VolumeSerialNumber").ToString()
    'End Function
    Friend Function GetMotherBoardID() As String
        If String.IsNullOrEmpty(Trim(strMotherBoardID)) Then
            Dim query As SelectQuery
            Dim search As ManagementObjectSearcher
            Dim info As ManagementObject
            Try
                query = New SelectQuery("Win32_BaseBoard")
                Try
                    search = New ManagementObjectSearcher(query)
                    For Each info In search.Get()
                        strMotherBoardID = info("SerialNumber").ToString()
                    Next
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            query = Nothing
            search = Nothing
            info = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return strMotherBoardID
    End Function
    Friend Function GetBIOSVersion() As String
        If String.IsNullOrEmpty(Trim(strBIOSVersion)) Then
            Dim query As SelectQuery
            Dim search As ManagementObjectSearcher
            Dim info As ManagementObject
            Try
                query = New SelectQuery("Win32_BIOS")
                Try
                    search = New ManagementObjectSearcher(query)
                    For Each info In search.Get()
                        strBIOSVersion = info("SerialNumber").ToString()
                    Next
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            query = Nothing
            search = Nothing
            info = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return strBIOSVersion
    End Function
    Friend Function GetBIOSSerialNumber() As String
        If String.IsNullOrEmpty(Trim(strBIOSSerialNumber)) Then
            Dim query As SelectQuery
            Dim search As ManagementObjectSearcher
            Dim info As ManagementObject
            Try
                query = New SelectQuery("Win32_BIOS")
                Try
                    search = New ManagementObjectSearcher(query)
                    For Each info In search.Get()
                        strBIOSSerialNumber = info("SMBIOSBIOSVersion").ToString()
                    Next
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
            query = Nothing
            search = Nothing
            info = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return strBIOSSerialNumber
    End Function
    Friend Sub saveSystemInfoInSettings()
        My.Settings.BIOSSerialNumber = GetBIOSSerialNumber()
        My.Settings.BIOSVersion = GetBIOSVersion()
        My.Settings.MACID = GetMACAddress()
        My.Settings.MotherboardID = GetMotherBoardID()
        My.Settings.ProcessorID = GetProcessorId()
        My.Settings.Save()
    End Sub
    Private Sub New()
        strProcessorID = GetProcessorId()
        strBIOSSerialNumber = GetBIOSSerialNumber()
        strBIOSVersion = GetBIOSVersion()
        strMotherBoardID = GetMotherBoardID()
        strMACAddress = GetMACAddress()
    End Sub
End Class
