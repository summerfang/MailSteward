Imports System.Collections.Generic
Imports Microsoft.Office.Interop.Outlook
Imports OutlookUtils

Public Class MainForm
    Private _olApp As Application = Nothing
    Private _olAddin As ThisAddIn = Nothing
    Private _olFoundFolder As Outlook.MAPIFolder = Nothing
    Private _oOA As OutlookAssistant = Nothing

    Public Sub New(ByVal olApp As Outlook.Application, ByVal olAddin As ThisAddIn, ByVal lFolders As List(Of Folder))
        InitializeComponent()

        _olApp = olApp
        _oOA = New OutlookAssistant(_olApp)
        _olAddin = olAddin
        '''ucPad.FullFolderList = lsAllFolderNames
        ucPad.AllFolders = lFolders
    End Sub

#Region "These codes handles pad's movement and size"
    'This variable is used for move the form
    Private mouseOffset As Point

    Private Sub HandleMouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim mousePos As Point = Control.MousePosition
            mousePos.Offset(mouseOffset.X, mouseOffset.Y)
            Location = mousePos
        End If
    End Sub

    Private Sub HandleMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If TypeOf (sender) Is Label Then
            mouseOffset = New Point(-e.X, -e.Y - lblMsg.Location.Y)
        Else
            mouseOffset = New Point(-e.X, -e.Y)
        End If
    End Sub

    Private Sub ucPad_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ucPad.MouseDown
        HandleMouseDown(sender, e)
    End Sub

    Private Sub ucPad_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ucPad.MouseMove
        HandleMouseMove(sender, e)
    End Sub

    Private Sub lblMsg_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMsg.MouseDown
        HandleMouseDown(sender, e)
    End Sub

    Private Sub lblMsg_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMsg.MouseMove
        HandleMouseMove(sender, e)
    End Sub

    Private Sub MainForm_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        HandleMouseDown(sender, e)
    End Sub

    Private Sub MainForm_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        HandleMouseMove(sender, e)
    End Sub

    'The MainForm's width varies following with the ucPad
    Private Sub ucPad_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ucPad.SizeChanged
        Me.Size = New Size(Math.Max(Math.Min(500, lblMsg.Width), (ucPad.Size.Width + 7 + btnClose.Width)), Me.Size.Height)
    End Sub
#End Region

    'It reset the _olAddin.frmDialog to nothing. It is used in _olAddin to judge will whether create the form again.
    Private Sub MainForm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        _olAddin.ResetMatchDialog()
    End Sub

    Private Sub MainForm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        'ucPad.txtCommandLine.Focus()
        ucPad.FocusCommandLine()
    End Sub

#Region "These codes handles moving the selected mails to specified folder. It is major part of the application function"
    Private Sub ucPad_FolderSelected() Handles ucPad.FolderSelected
        MoveCurrentMailTo(ucPad.SelectedFolder)
    End Sub

    Sub MoveCurrentMailTo(ByVal oFolder As Folder)
        On Error GoTo MoveMail_Error

        'Step 1. Protection code
        If oFolder Is Nothing Then
            lblMsg.Text = "No such folder!"
            lblMsg.ForeColor = Color.Red
            Return
        End If

        'Step 2. Get a folder's full path even the name contains "\" or "/"
        Dim sFolderPath As String
        sFolderPath = Microsoft.VisualBasic.Left(oFolder.FolderPath, InStrRev(oFolder.FolderPath, "\")) + oFolder.Name

        'Step 3. Get selection.
        Dim myOlExp As Outlook.Explorer
        Dim myOlSel As Outlook.Selection

        myOlExp = _olApp.ActiveExplorer
        myOlSel = myOlExp.Selection

        'Step 4. Avoid crash if nothing is selected.
        Dim iSelCount As Integer
        iSelCount = myOlSel.Count
        If iSelCount <= 0 Then
            Return
        End If

        'Step 5. Move selected items to folder
        Dim sSubject As String = ""

        Dim i As Integer

        For i = 1 To iSelCount
            Dim o As Object
            o = myOlSel.Item(i)
            If TypeOf o Is Outlook.MailItem Then
                Dim mailItem As Outlook.MailItem = TryCast(o, MailItem)
                sSubject = mailItem.Subject
                mailItem.Move(oFolder)
            ElseIf (TypeOf o Is Outlook.ContactItem) Then
                Dim contactItem As Outlook.ContactItem = TryCast(o, Outlook.ContactItem)
                sSubject = contactItem.Subject
                contactItem.Move(oFolder)
            ElseIf (TypeOf o Is Outlook.AppointmentItem) Then
                Dim apptItem As Outlook.AppointmentItem = TryCast(o, Outlook.AppointmentItem)
                sSubject = apptItem.Subject
                apptItem.Move(oFolder)
            ElseIf (TypeOf o Is Outlook.TaskItem) Then
                Dim taskItem As Outlook.TaskItem = TryCast(o, Outlook.TaskItem)
                sSubject = taskItem.Subject
                taskItem.Move(oFolder)
            ElseIf (TypeOf o Is Outlook.MeetingItem) Then
                Dim meetingItem As Outlook.MeetingItem = TryCast(o, Outlook.MeetingItem)
                sSubject = meetingItem.Subject
                meetingItem.Move(oFolder)
            End If
        Next i

        'Step 6.Display success message.
        If iSelCount > 1 Then
            lblMsg.Text = "The " + Str(i) + " items including " + sSubject + " are moved to " + sFolderPath
        Else
            lblMsg.Text = sSubject + " is moved to " + sFolderPath
        End If

        lblMsg.ForeColor = Color.Blue

        Return

        'If anything wrong, display warning message.
MoveMail_Error:
        lblMsg.Text = "Unable to move the any item to " + sFolderPath
        lblMsg.ForeColor = Color.Red
    End Sub
#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class