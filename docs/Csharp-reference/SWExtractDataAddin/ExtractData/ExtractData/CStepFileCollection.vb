Option Explicit On
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Friend Class CStepFileCollection
    Implements INotifyPropertyChanged

    Private oStepFileCollection As ObservableCollection(Of CStepFile)
    Private isProfileDataPopulated As Boolean = False
    Private isProfileDataExtracted As Boolean = False
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub New()
        oStepFileCollection = New ObservableCollection(Of CStepFile)
    End Sub
    Public Property STEPFileCollection() As ObservableCollection(Of CStepFile)
        Get
            Return oStepFileCollection
        End Get
        Set(value As ObservableCollection(Of CStepFile))
            oStepFileCollection = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Sub Clear()
        On Error Resume Next
        If Not oStepFileCollection Is Nothing Then
            oStepFileCollection.Clear()
            IsDataPopulated = False
            IsDataExtracted = False
        End If
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub Add(ByRef oStepFile As CStepFile)
        On Error Resume Next
        If Not oStepFileCollection Is Nothing And Not oStepFile Is Nothing Then
            If DoesFileWithThisNameAlreadyExist(oStepFile.FullFileName, -1) = False Then
                oStepFileCollection.Add(oStepFile)
                If Count > 0 Then
                    IsDataPopulated = True
                End If
            End If
        End If
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public ReadOnly Property Count() As Integer
        Get
            If Not oStepFileCollection Is Nothing Then
                Return oStepFileCollection.Count
            End If
            Return 0
        End Get
    End Property
    Public Property IsDataPopulated As Boolean
        Get
            Return isProfileDataExtracted
        End Get
        Set(value As Boolean)
            isProfileDataExtracted = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property IsDataExtracted As Boolean
        Get
            Return isProfileDataPopulated
        End Get
        Set(value As Boolean)
            isProfileDataPopulated = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property IsShadedWithEdges As Boolean
        Get
            Return My.Settings.ShadedWithEdges
        End Get
        Set(value As Boolean)
            My.Settings.ShadedWithEdges = value
            My.Settings.Save()
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Sub UpdateCheckStatus(ByRef isCheck As Boolean)
        On Error Resume Next
        If Not oStepFileCollection Is Nothing Then
            Dim i As Integer = 0
            For i = 0 To Count - 1
                oStepFileCollection.Item(i).IsSelected = isCheck
            Next i
        End If
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function DoesFileWithThisNameAlreadyExist(ByRef strFileFullName As String, Optional ByRef iIndexOfFile As Integer = -1) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oStepFileCollection Is Nothing And Not String.IsNullOrEmpty(strFileFullName) Then
            Dim i As Integer = 0
            While i < oStepFileCollection.Count And bReturn = False
                If oStepFileCollection.Item(i).FullFileName = strFileFullName Then
                    bReturn = True
                    iIndexOfFile = i
                End If
                i = i + 1
            End While
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Sub ExtractData()
        On Error Resume Next
        If Not oStepFileCollection Is Nothing Then
            Dim i As Integer = 0
            CProgress.GetInstance.Minimum = 0
            CProgress.GetInstance.Maximum = Count
            CProgress.GetInstance.ProgressText = "Extracting Data..."
            CProgress.GetInstance.InitialiseProgressLevel()
            CProgress.GetInstance.ProgressBarVisibility = Visibility.Visible
            System.Windows.Forms.Application.DoEvents()
            For i = 0 To Count - 1
                oStepFileCollection.Item(i).ExtractData()
                CProgress.GetInstance.ProgressText = "Extracting Data (" & i + 1 & "/" & CProgress.GetInstance().Maximum & ")..."
                CProgress.GetInstance.UpdateValue()
            Next i
            CProgress.GetInstance.Minimum = 0
            CProgress.GetInstance.Maximum = 100
            CProgress.GetInstance.ProgressText = ""
            CProgress.GetInstance.InitialiseProgressLevel()
            CProgress.GetInstance.ProgressBarVisibility = Visibility.Collapsed
            System.Windows.Forms.Application.DoEvents()
            IsDataExtracted = True
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub SaveDataAndFile()
        On Error Resume Next
        If Not oStepFileCollection Is Nothing Then
            CProgress.GetInstance.Minimum = 0
            CProgress.GetInstance.Maximum = Count
            CProgress.GetInstance.ProgressText = "Saving data and files..."
            CProgress.GetInstance.InitialiseProgressLevel()
            CProgress.GetInstance.ProgressBarVisibility = Visibility.Visible
            System.Windows.Forms.Application.DoEvents()
            Dim i As Integer = 0
            For i = 0 To Count - 1
                oStepFileCollection.Item(i).SaveDataAsCustomProperties()
                oStepFileCollection.Item(i).SaveModel()
                CProgress.GetInstance.ProgressText = "Saving data (" & i + 1 & "/" & CProgress.GetInstance().Maximum & ")..."
                CProgress.GetInstance.UpdateValue()
            Next i
            CProgress.GetInstance.Minimum = 0
            CProgress.GetInstance.Maximum = 100
            CProgress.GetInstance.ProgressText = ""
            CProgress.GetInstance.InitialiseProgressLevel()
            CProgress.GetInstance.ProgressBarVisibility = Visibility.Collapsed
            System.Windows.Forms.Application.DoEvents()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub UpdateView()
        On Error Resume Next
        If Not oStepFileCollection Is Nothing Then
            For i As Integer = 0 To oStepFileCollection.Count - 1
                oStepFileCollection.Item(i).UpdateView()
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        Clear()
        oStepFileCollection = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
