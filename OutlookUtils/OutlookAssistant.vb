Imports Microsoft.Office.Interop.Outlook

Public Class OutlookAssistant
    Private _oOutlookApp As Application = Nothing

    'Return all folder objects in a list in outlook
    Public Function GetAllFoldersInStores() As List(Of Folder)
        'Dim olApp As New Outlook.Application
        Dim colStores As Stores
        Dim oStore As Store
        Dim oRoot As Folder
        Dim lFolderR As New List(Of Folder)

        On Error Resume Next
        colStores = _oOutlookApp.Session.Stores
        For Each oStore In colStores
            oRoot = oStore.GetRootFolder
            lFolderR.Add(oRoot)
            EnumerateAllFolders(oRoot, lFolderR)
        Next

        Return lFolderR
    End Function

    Private Sub EnumerateAllFolders(ByVal oFolder As Folder, ByRef lFolder As List(Of Folder))
        Dim folders As Folders
        Dim Folder As Folder
        Dim foldercount As Integer

        On Error Resume Next
        folders = oFolder.Folders
        foldercount = folders.Count
        'Check if there are any folders below oFolder
        If foldercount Then
            For Each Folder In folders
                lFolder.Add(Folder)
                EnumerateAllFolders(Folder, lFolder)
            Next
        End If
    End Sub

    Function MoveCurrentMailsToFolder(ByVal oFolder As Folder) As Boolean
        If oFolder Is Nothing Then
            Return False
        End If

        Dim myOlExp As Explorer
        Dim myOlSel As Selection

        myOlExp = _oOutlookApp.ActiveExplorer
        myOlSel = myOlExp.Selection

        Dim currentMailItem As MailItem = Nothing

        Dim i As Integer

        For i = 1 To myOlSel.Count
            currentMailItem = myOlSel.Item(i)
            currentMailItem.Move(oFolder)
        Next i
        Return True

MoveFolder_Error:
        MoveCurrentMailsToFolder = False
        Exit Function
    End Function

    Public Sub New(ByVal oOutlookApp As Application)
        _oOutlookApp = oOutlookApp
    End Sub
End Class
