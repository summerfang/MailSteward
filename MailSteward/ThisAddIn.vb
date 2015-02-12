Imports OutlookUtils
Imports System.Collections.Generic

Public Class ThisAddIn
    Public WithEvents _olExplorer As Outlook.Explorer

    Private _olHelperToolbar As Office.CommandBar
    Private _btnShowHelper As Office.CommandBarButton
    Private _olSelectExplorers As Outlook.Explorers

    Private _oOA As OutlookAssistant 'The object is used to operate outlook.

    Private _lAllFolders As List(Of Outlook.Folder)
    Private _lAllFolderCollection As New List(Of Outlook.Folders)

    'It is ime like pad used to operate
    Private _frmMatch As MainForm


    Public Sub ResetMatchDialog()
        _frmMatch = Nothing

    End Sub

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        'Step 1. Create the toolbar and button
        _olExplorer = Me.Application.ActiveExplorer
        _olSelectExplorers = Me.Application.Explorers()

        AddHandler _olSelectExplorers.NewExplorer, AddressOf Me.NewExplorer_Event
        AddToolbar()

        'Step 2. Create the _oOA for outlook operation
        _oOA = New OutlookAssistant(Me.Application)

        'Step 3. Display the pad after the addin starts
        _lAllFolders = _oOA.GetAllFoldersInStores()

        Dim i As Integer
        For i = 0 To _lAllFolders.Count - 1
            Dim oFolderCollection As Outlook.Folders

            oFolderCollection = _lAllFolders(i).Folders
            _lAllFolderCollection.Add(oFolderCollection)

            AddHandler _lAllFolderCollection(i).FolderAdd, AddressOf Me.FoldersEvent_FolderAdd
            AddHandler _lAllFolderCollection(i).FolderChange, AddressOf Me.FoldersEvent_FolderChange
            AddHandler _lAllFolderCollection(i).FolderRemove, AddressOf Me.FoldersEvent_FolderRemove
        Next

        _frmMatch = New MainForm(Me.Application, Me, _lAllFolders)
        _frmMatch.Show()

    End Sub

    Private Sub _olExplorer_SelectionChange() Handles _olExplorer.SelectionChange
        'Return if nothing is selected.
        If _olExplorer.Selection.Count <= 0 Then
            Return
        End If

        Dim sSubject As String
        Dim selObject As Object = Me.Application.ActiveExplorer.Selection.Item(1)

        If (TypeOf selObject Is Outlook.MailItem) Then
            Dim mailItem As Outlook.MailItem = TryCast(selObject, Outlook.MailItem)
            sSubject = mailItem.Subject + " " + mailItem.SenderName
            _frmMatch.ucPad.ProcessMail(sSubject)

        ElseIf (TypeOf selObject Is Outlook.ContactItem) Then
            Dim contactItem As Outlook.ContactItem = TryCast(selObject, Outlook.ContactItem)
            sSubject = contactItem.Subject + " " + contactItem.FullName
            _frmMatch.ucPad.ProcessMail(sSubject)

        ElseIf (TypeOf selObject Is Outlook.AppointmentItem) Then
            Dim apptItem As Outlook.AppointmentItem = TryCast(selObject, Outlook.AppointmentItem)
            sSubject = apptItem.Subject + " " + apptItem.Organizer
            _frmMatch.ucPad.ProcessMail(sSubject)

        ElseIf (TypeOf selObject Is Outlook.TaskItem) Then
            Dim taskItem As Outlook.TaskItem = TryCast(selObject, Outlook.TaskItem)
            sSubject = taskItem.Subject + " " + taskItem.Owner
            _frmMatch.ucPad.ProcessMail(sSubject)

        ElseIf (TypeOf selObject Is Outlook.MeetingItem) Then
            Dim meetingItem As Outlook.MeetingItem = TryCast(selObject, Outlook.MeetingItem)
            sSubject = meetingItem.Subject + " " + meetingItem.SenderName
            _frmMatch.ucPad.ProcessMail(sSubject)
        Else
            _frmMatch.ucPad.ProcessMail("")
        End If
    End Sub

    Private Sub NewExplorer_Event(ByVal new_Explorer As Outlook.Explorer)
        new_Explorer.Activate()
        _olHelperToolbar = Nothing
        Call Me.AddToolbar()
    End Sub

    Private Sub AddToolbar()
        Dim btn As Office.CommandBarButton
        If _olHelperToolbar Is Nothing Then
            Dim cmdBars As Office.CommandBars = Me.Application.ActiveExplorer().CommandBars
            _olHelperToolbar = cmdBars.Add("OutlookHelper", Office.MsoBarPosition.msoBarTop, False, True)
        End If
        Try
            btn = CType(_olHelperToolbar.Controls.Add(1), Office.CommandBarButton)
            With btn
                .Style = Office.MsoButtonStyle.msoButtonCaption
                .Caption = "Show helper"
                .Tag = "Show Helper"
            End With
            If Me._btnShowHelper Is Nothing Then
                Me._btnShowHelper = btn
                AddHandler _btnShowHelper.Click, AddressOf ButtonMailPadClick
            End If
            _olHelperToolbar.Visible = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ButtonMailPadClick(ByVal ctrl As Office.CommandBarButton, ByRef Cancel As Boolean)
        If _frmMatch Is Nothing Then
            _frmMatch = New MainForm(Me.Application, Me, _lAllFolders)
        End If
        _frmMatch.Show()
        _frmMatch.BringToFront()
    End Sub

    Private Sub FoldersEvent_FolderAdd(ByVal Folder As Microsoft.Office.Interop.Outlook.MAPIFolder)
        Debug.Print("FoldersEvent_FolderAdd" + Folder.FolderPath)
        _lAllFolders.Add(Folder)
        _frmMatch.ucPad.AllFolders = _lAllFolders

        Dim oFolders As Outlook.Folders
        oFolders = Folder.Folders

        AddHandlerToFolders(oFolders)

        _lAllFolderCollection.Add(oFolders)
    End Sub

    Private Sub FoldersEvent_FolderChange(ByVal Folder As Microsoft.Office.Interop.Outlook.MAPIFolder)
        Debug.Print("FoldersEvent_FolderChange" & Folder.FolderPath & " ID=" & CStr(Folder.EntryID) & " HashCode=" & CStr(Folder.GetHashCode))

        Dim iR As Integer
        iR = FindFolder(Folder)

        If -1 = iR Then
            Debug.Print("Can't find the folder " & Folder.FolderPath)
        Else
            '            _lAllFolders(iR).Name = Folder.Name
            _lAllFolders.RemoveAt(iR)
            _lAllFolders.Insert(iR, Folder)
            Debug.Print("_lAllFolders(iR).Name:" + _lAllFolders(iR).Name)
            _frmMatch.ucPad.AllFolders = _lAllFolders

            Dim sSubject As String
            sSubject = GetActiveSubject()
            _frmMatch.ucPad.ProcessMail(sSubject)  'Avoid orginal selection is still here
        End If

        'RefreshFolder()
    End Sub

    Private Sub FoldersEvent_FolderRemove()
        Debug.Print("FoldersEvent_FolderRemove")
        RefreshFolder()
    End Sub

    Private Sub RefreshFolder()
        RemoveFoldersCollectionEvents(_lAllFolderCollection)
        _lAllFolderCollection.Clear()
        _lAllFolders.Clear()

        _lAllFolders = _oOA.GetAllFoldersInStores
        _frmMatch.ucPad.AllFolders = _lAllFolders

        Dim sSubject As String
        sSubject = GetActiveSubject()
        _frmMatch.ucPad.ProcessMail(sSubject)  'Avoid orginal selection is still here

        Dim i As Integer

        For i = 0 To _lAllFolders.Count - 1
            Dim oFolderCollection As Outlook.Folders

            oFolderCollection = _lAllFolders(i).Folders
            _lAllFolderCollection.Add(oFolderCollection)

            AddHandler _lAllFolderCollection(i).FolderAdd, AddressOf Me.FoldersEvent_FolderAdd
            AddHandler _lAllFolderCollection(i).FolderChange, AddressOf Me.FoldersEvent_FolderChange
            AddHandler _lAllFolderCollection(i).FolderRemove, AddressOf Me.FoldersEvent_FolderRemove
        Next
    End Sub

    'Remove all events handlers in a list of folders
    Private Sub RemoveFoldersCollectionEvents(ByRef lFolders As List(Of Outlook.Folders))
        On Error GoTo RemoveFolderEvents_Error
        Dim i As Integer
        For i = 0 To lFolders.Count - 1
            RemoveHandler lFolders(i).FolderAdd, AddressOf FoldersEvent_FolderAdd
            RemoveHandler lFolders(i).FolderChange, AddressOf FoldersEvent_FolderChange
            RemoveHandler lFolders(i).FolderRemove, AddressOf FoldersEvent_FolderRemove
        Next
        Exit Sub

RemoveFolderEvents_Error:
        Debug.Print("RemoveFolderEvents failed")
    End Sub

    Private Sub AddHandlerToFolders(ByRef oFolders As Outlook.Folders)
        AddHandler oFolders.FolderAdd, AddressOf FoldersEvent_FolderAdd
        AddHandler oFolders.FolderChange, AddressOf FoldersEvent_FolderChange
        AddHandler oFolders.FolderRemove, AddressOf FoldersEvent_FolderRemove
    End Sub

    Private Function FindFolder(ByVal oFolder As Outlook.Folder) As Integer
        Dim i, iR As Integer

        iR = -1
        For i = 0 To _lAllFolders.Count - 1
            If (oFolder.EntryID = _lAllFolders(i).EntryID) And (oFolder.Name <> _lAllFolders(i).Name) Then
                iR = i
                Exit For
            End If
        Next

        Return iR
    End Function

    Private Function GetActiveSubject() As String
        Dim olExporer As Outlook.Explorer

        olExporer = Me.Application.ActiveExplorer

        'Return if nothing is selected.
        If olExporer.Selection.Count <= 0 Then
            Return ""
        End If

        Dim sSubject As String
        Dim selObject As Object = Me.Application.ActiveExplorer.Selection.Item(1)

        If (TypeOf selObject Is Outlook.MailItem) Then
            Dim mailItem As Outlook.MailItem = TryCast(selObject, Outlook.MailItem)
            sSubject = mailItem.Subject + " " + mailItem.SenderName

        ElseIf (TypeOf selObject Is Outlook.ContactItem) Then
            Dim contactItem As Outlook.ContactItem = TryCast(selObject, Outlook.ContactItem)
            sSubject = contactItem.Subject + " " + contactItem.FullName
            _frmMatch.ucPad.ProcessMail(sSubject)

        ElseIf (TypeOf selObject Is Outlook.AppointmentItem) Then
            Dim apptItem As Outlook.AppointmentItem = TryCast(selObject, Outlook.AppointmentItem)
            sSubject = apptItem.Subject + " " + apptItem.Organizer

        ElseIf (TypeOf selObject Is Outlook.TaskItem) Then
            Dim taskItem As Outlook.TaskItem = TryCast(selObject, Outlook.TaskItem)
            sSubject = taskItem.Subject + " " + taskItem.Owner

        ElseIf (TypeOf selObject Is Outlook.MeetingItem) Then
            Dim meetingItem As Outlook.MeetingItem = TryCast(selObject, Outlook.MeetingItem)
            sSubject = meetingItem.Subject + " " + meetingItem.SenderName
        Else
            Return ""
        End If

        Return sSubject
    End Function
End Class
