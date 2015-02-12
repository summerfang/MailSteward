Imports System.Collections.Generic
Imports Microsoft.Office.Interop

Public Class ucPad
    Private _lMatchedFolders As List(Of Outlook.Folder) = New List(Of Outlook.Folder)
    Private _oDFD = New DisplayFoldersData
    Public Event FolderSelected()

    Private Sub txtCommandLine_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCommandLine.KeyUp
        Select Case e.KeyValue
            Case Keys.Left, Keys.Up
                dfbFolder.UpdateViewForPreviousFolder()

            Case Keys.Right, Keys.Down
                dfbFolder.UpdateViewForNextFolder()

            Case Keys.Return
                If _lMatchedFolders.Count > 0 Then
                    _oSelectedFolder = dfbFolder.GetDisplayFoldersData.GetFolders(dfbFolder.GetDisplayFoldersData.GetSelectFolder - 1)
                    RaiseEvent FolderSelected()
                End If
        End Select
    End Sub

    Private Sub txtCommandLine_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCommandLine.TextChanged
        Dim s As String
        s = Trim(txtCommandLine.Text)

        Dim lsSubjects As New List(Of String)
        lsSubjects.Add(s)

        Dim lMatchedFolders As List(Of Outlook.Folder)
        lMatchedFolders = GetMatchedFolders(lsSubjects)
        _lMatchedFolders = lMatchedFolders

        _oDFD = New DisplayFoldersData(_lMatchedFolders)
        dfbFolder.SetDisplayFoldersData(_oDFD)
        dfbFolder.UpdateAllFolders()
    End Sub

    'It handles the mouse doubleclick event of selected folder
    Private Sub DFB_FolderSelected() Handles dfbFolder.FolderSelected
        If _lMatchedFolders.Count > 0 Then
            _oSelectedFolder = _lMatchedFolders(dfbFolder.GetDisplayFoldersData.GetSelectFolder - 1)
        End If
        RaiseEvent FolderSelected()
    End Sub

    Private Function GetMatchedFolders(ByVal lsSubjects As List(Of String)) As List(Of Outlook.Folder)
        Dim lFoldersR As New List(Of Outlook.Folder)

        For Each oFolder As Outlook.Folder In _lAllFolders
            Dim iMatchedLevel As Integer
            iMatchedLevel = IsFolderMatchOneSubject(oFolder, lsSubjects)

            If iMatchedLevel = 2 Then
                lFoldersR.Insert(0, oFolder)
            ElseIf iMatchedLevel = 1 Then
                lFoldersR.Add(oFolder)
            End If
        Next

        Return lFoldersR
    End Function

    ''''Check whether the folder name matches any subject.
    Private Function IsFolderMatchOneSubject(ByVal oFolder As Outlook.Folder, ByVal lsSubjects As List(Of String)) As Integer
        Dim iR As Integer = 0

        Dim sFolder As String
        sFolder = oFolder.Name

        For Each s As String In lsSubjects
            Dim iMatchedLevel As Integer
            iMatchedLevel = IsMatchedString(sFolder, s)
            If iMatchedLevel Then
                iR = iMatchedLevel
                Exit For
            End If
        Next

        Return iR
    End Function

    Private Function IsMatchedString(ByVal strSubjectWord As String, ByVal strFolderName As String) As Integer
        Dim iMatchedLevel As Integer = 0

        If Trim(strSubjectWord) <> "" And strFolderName <> "" Then
            If (LCase(Trim(strFolderName)) = LCase(Trim(strSubjectWord))) Then
                iMatchedLevel = 2
            ElseIf (InStr(LCase(Trim(strSubjectWord)), LCase(Trim(strFolderName))) <> 0) Or (InStr(LCase(Trim(strFolderName)), LCase(Trim(strSubjectWord))) <> 0) Then
                iMatchedLevel = 1
            End If
        End If
        Return iMatchedLevel
    End Function

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub

#Region "Customized properties of ucPad"
    'It stores folder object. Need container to pass the value
    Private _lAllFolders As New List(Of Outlook.Folder)

    Public WriteOnly Property AllFolders() As List(Of Outlook.Folder)
        'Get
        '    Return lFullFolderList
        'End Get
        Set(ByVal value As List(Of Outlook.Folder))
            _lAllFolders = value
            Me.Invalidate()
        End Set
    End Property

    'It store the selected folder's name
    Private _oSelectedFolder As Outlook.Folder

    Public ReadOnly Property SelectedFolder() As Outlook.Folder
        Get
            Return _oSelectedFolder
        End Get
    End Property

#End Region

    Public Sub ProcessMail(ByVal sSubject As String)
        Dim lsSubjects As New List(Of String)
        Dim astr() As String
        astr = Split(Trim(sSubject))
        For Each s In astr
            If Trim(s) <> "" Then
                lsSubjects.Add(s)
            End If
        Next

        Dim lsMatchedFolderPath As List(Of Outlook.Folder)
        lsMatchedFolderPath = GetMatchedFolders(lsSubjects)
        _lMatchedFolders = lsMatchedFolderPath

        _oDFD = New DisplayFoldersData(_lMatchedFolders)
        dfbFolder.SetDisplayFoldersData(_oDFD)
        dfbFolder.UpdateAllFolders()

        txtCommandLine.SelectAll()
    End Sub

    Public Sub FocusCommandLine()
        txtCommandLine.Focus()
    End Sub

    Private Sub dfbFolder_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dfbFolder.SizeChanged
        Me.Size = New Size(dfbFolder.Width, Me.Height)
    End Sub
End Class
