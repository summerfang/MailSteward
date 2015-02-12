Imports Microsoft.Office.Interop

Public Class DisplayFoldersBar
    Private _oDFD As DisplayFoldersData = New DisplayFoldersData

    Public Event FolderSelected()

    Sub SetDisplayFoldersData(ByVal oDFD)
        _oDFD = oDFD
    End Sub

    Function GetDisplayFoldersData() As DisplayFoldersData
        Return _oDFD
    End Function

    Sub UpdateViewForPreviousFolder()
        _oDFD.PreviousFolder()
        UpdateAllFolders()
    End Sub

    Sub UpdateViewForNextFolder()
        _oDFD.NextFolder()
        UpdateAllFolders()
    End Sub

    Sub UpdateAllFolders()
        Dim iStartPositionX As Integer = 0
        Dim iStartPositionY As Integer = 4

        Controls.Clear()

        'Step 1. If there are no folder is found to match the mail, display in red.
        If _oDFD.GetFolders.Count = 0 Then
            'Generate order labels
            Dim lNotFound As New Label

            With lNotFound
                .Name = "lblNotFound"
                .Text = "No folder is found to match the mail!"
                .ForeColor = Color.Red
                .Location = New Point(iStartPositionX, 0)
                .Size = New Size(.PreferredWidth, .PreferredHeight)
                iStartPositionX += .Size.Width
            End With
            Controls.Add(lNotFound)
        End If

        'Step 2. Draw previous button
        Dim lblPreviousFolder As New Label
        If _oDFD.GetFirstDisplayFolders > 1 Then
            With lblPreviousFolder
                .Name = "lblPreviousFolder"
                '.Text = "<<"
                .Image = imgNavigator.Images(0)
                .Location = New Point(iStartPositionX, 0)
                .Size = New Size(21, 21)
                iStartPositionX += .Size.Width
                AddHandler .Click, AddressOf Label_ClickPreviousFolder
            End With
            Controls.Add(lblPreviousFolder)
        End If

        'Step 3. Draw the next five label
        For i As Integer = 1 To Math.Min(5, _oDFD.GetFolders.Count)

            'Generate order labels
            Dim lOrderNum As New Label

            With lOrderNum
                .Name = "O" + i.ToString
                .Text = (_oDFD.GetFirstDisplayFolders + i - 1).ToString + "."
                .Location = New Point(iStartPositionX, iStartPositionY)
                .Size = New Size(.PreferredWidth, .PreferredHeight)
                iStartPositionX += .Size.Width
            End With

            Controls.Add(lOrderNum)

            'Generate labels that contains matched folder names.
            Dim newLabel As New Label

            With newLabel
                .Name = "l" + i.ToString
                .UseMnemonic = False

                If HasSameNameFolder(_oDFD.GetFolders.Item(_oDFD.GetFirstDisplayFolders + i - 2), _oDFD.GetFolders) Then
                    .Text = _oDFD.GetFolders.Item(_oDFD.GetFirstDisplayFolders + i - 2).FolderPath
                Else
                    .Text = _oDFD.GetFolders.Item(_oDFD.GetFirstDisplayFolders + i - 2).Name
                End If


                If i = _oDFD.GetPoint Then
                    .ForeColor = Color.White
                    .BackColor = Color.Blue
                End If

                .Location = New Point(iStartPositionX, iStartPositionY)
                .Size = New Size(.PreferredWidth, .PreferredHeight)
                iStartPositionX += newLabel.Size.Width

                AddHandler .DoubleClick, AddressOf Label_DoubleClickSelectFolder
            End With

            Controls.Add(newLabel)
        Next

        'Step 4. Draw the next button
        Dim lblNextFolder As New Label

        If _oDFD.GetFolders.Count > 5 And _oDFD.GetFolders.Count - _oDFD.GetFirstDisplayFolders > 4 Then
            With lblNextFolder
                .Name = "btnNextFolder"
                .Image = imgNavigator.Images(1)

                .Location = New Point(iStartPositionX, 0)
                .Size = New Size(21, 21)
                iStartPositionX += .Size.Width
                AddHandler .Click, AddressOf Label_ClickNextFolder

            End With
            Controls.Add(lblNextFolder)
        End If

        Me.Size = New Size(iStartPositionX + 10, Me.Height)
    End Sub

    'A list contains two same string means the folder name is same.
    Private Function HasSameNameFolder(ByVal oAFolder As Outlook.Folder, ByVal ls As List(Of Outlook.Folder)) As Boolean
        Dim b As Boolean = False
        Dim iCount As Integer = 0
        Dim sFolder As String

        For Each o As Outlook.Folder In ls
            sFolder = o.Name
            If LCase(Trim(sFolder)) = LCase(Trim(oAFolder.Name)) Then
                iCount += 1
            End If
        Next

        If iCount > 1 Then
            b = True
        Else
            b = False
        End If

        Return b
    End Function

    Private Sub Label_DoubleClickSelectFolder(ByVal sender As Object, ByVal e As EventArgs)
        If TypeOf sender Is Label Then
            Dim lLabel As Label
            lLabel = CType(sender, Label)

            Dim iSelected As Integer
            iSelected = Val(Microsoft.VisualBasic.Right(lLabel.Name, lLabel.Name.Length - 1))

            Dim iPoint As Integer
            iPoint = _oDFD.GetPoint()

            If (iPoint < iSelected) Then
                For i = iPoint To iSelected - 1
                    _oDFD.NextFolder()
                Next
            ElseIf (iPoint > iSelected + 1) Then
                For i = iPoint To iSelected Step -1
                    _oDFD.PreviousFolder()
                Next
            End If

            RaiseEvent FolderSelected()
        End If
    End Sub

    Private Sub Label_ClickPreviousFolder(ByVal sender As Object, ByVal e As EventArgs)
        UpdateViewForPreviousFolder()
    End Sub

    Private Sub Label_ClickNextFolder(ByVal sender As Object, ByVal e As EventArgs)
        UpdateViewForNextFolder()
    End Sub

End Class
