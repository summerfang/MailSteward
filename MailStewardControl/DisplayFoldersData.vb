Imports Microsoft.Office.Interop

Public Class DisplayFoldersData

    Private _iPoint As Integer
    Private _iFirstDisplayFolders As Integer
    Private _iSelectedFolder As Integer
    Private _lFolders As List(Of Outlook.Folder)

    Private _oDFB As DisplayFoldersBar

    Public Sub New()
        _iPoint = 0
        _iFirstDisplayFolders = 0
        _iSelectedFolder = 0
        _lFolders = New List(Of Outlook.Folder)
    End Sub

    Public Sub New(ByVal lFolders As List(Of Outlook.Folder))
        _lFolders = lFolders
        If _lFolders.Count = 0 Then
            _iPoint = 0
            _iFirstDisplayFolders = 0
            _iSelectedFolder = 0
        Else
            _iPoint = 1
            _iFirstDisplayFolders = 1
            _iSelectedFolder = 1
        End If

        Debug.Print("_iPoint=" + CStr(_iPoint) + ";_iFirstDisplayFolders=" + CStr(_iFirstDisplayFolders) + ";_iSelectedFolder" + CStr(_iSelectedFolder))
    End Sub

    Public Sub PreviousFolder()
        If _iPoint > 1 Then
            _iPoint -= 1
        End If

        If _iSelectedFolder > 1 Then
            _iSelectedFolder -= 1
        End If

        If _iPoint = 1 And _iFirstDisplayFolders > _iSelectedFolder Then
            _iFirstDisplayFolders -= 1
        End If
        Debug.Print("_iPoint=" + CStr(_iPoint) + ";_iFirstDisplayFolders=" + CStr(_iFirstDisplayFolders) + ";_iSelectedFolder" + CStr(_iSelectedFolder))

    End Sub

    Public Sub NextFolder()
        If _iPoint < Math.Min(_lFolders.Count, 5) Then
            _iPoint += 1
        End If

        If _iSelectedFolder < _lFolders.Count Then
            _iSelectedFolder += 1
        End If

        If _iPoint = 5 And _iSelectedFolder - _iFirstDisplayFolders > 4 Then
            _iFirstDisplayFolders += 1
        End If
        Debug.Print("_iPoint=" + CStr(_iPoint) + ";_iFirstDisplayFolders=" + CStr(_iFirstDisplayFolders) + ";_iSelectedFolder" + CStr(_iSelectedFolder))
    End Sub

    Public Function GetFolders() As List(Of Outlook.Folder)
        Return _lFolders
    End Function

    Public Function GetPoint() As Integer
        Return _iPoint
    End Function

    Public Function GetFirstDisplayFolders() As Integer
        Return _iFirstDisplayFolders
    End Function

    Public Function GetSelectFolder() As Integer
        Return _iSelectedFolder
    End Function

End Class
