Option Explicit On
Friend Class CProgress
    Implements System.ComponentModel.INotifyPropertyChanged

    Private strProgressText As String = ""
    Private iProgressLevel As Integer = 0
    Private iMinimumLevel As Integer = 0
    Private iMaximumLevel As Integer = 100
    Private oVisibilityOfControls As Visibility = Visibility.Hidden
    Private isProgressBarIndeterminate As Boolean = False

    Private Shared myInstance As CProgress
    Public Shared Function GetInstance() As CProgress
        If myInstance Is Nothing Then
            myInstance = New CProgress()
        End If
        Return myInstance
    End Function
    Private Sub New()
        strProgressText = ""
        iProgressLevel = 0
        oVisibilityOfControls = Visibility.Hidden
        isProgressBarIndeterminate = False
    End Sub
    Public Property ProgressText As String
        Get
            Return strProgressText
        End Get
        Set(value As String)
            strProgressText = value
            If String.IsNullOrEmpty(strProgressText) Then
                ProgressBarVisibility = Visibility.Hidden
            Else
                ProgressBarVisibility = Visibility.Visible
            End If
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(""))
        End Set
    End Property
    Public ReadOnly Property ProgressLevel As Integer
        Get
            Return iProgressLevel
        End Get
    End Property
    Public Property ProgressBarVisibility As Visibility
        Get
            Return oVisibilityOfControls
        End Get
        Set(value As Visibility)
            oVisibilityOfControls = value
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property IsIndeterminate As Boolean
        Get
            Return isProgressBarIndeterminate
        End Get
        Set(value As Boolean)
            isProgressBarIndeterminate = value
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property Minimum As Integer
        Get
            Return iMinimumLevel
        End Get
        Set(value As Integer)
            iMinimumLevel = value
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property Maximum As Integer
        Get
            Return iMaximumLevel
        End Get
        Set(value As Integer)
            iMaximumLevel = value
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Sub InitialiseProgressLevel()
        iProgressLevel = 0
    End Sub
    Public Sub UpdateValue()
        On Error Resume Next
        If iProgressLevel < iMaximumLevel Then
            iProgressLevel = iProgressLevel + 1
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(""))
            System.Windows.Forms.Application.DoEvents()
        Else
            iProgressLevel = Maximum
            RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(""))
            System.Windows.Forms.Application.DoEvents()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

End Class
