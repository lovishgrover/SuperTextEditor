'Author : Lovish Grover
'Date : March 24 2020
'Description :
' This code is for text editor in which a user can write text and perform functions on 
' that text like save file,copy,cut,paste ,opening a file and also create a new file.
' Some of the code is taken by the resources provided by Prof. Kyle Chapman 

'Github link =


Option Strict On

Imports System.IO
Public Class frmEditor
    ' Declarations

    Dim isFileSave As Boolean = False
    Dim openFilePath As String = String.Empty 'to open a file with the path selected
    Dim saveFilePath As String = String.Empty 'to store the path to save location of file

#Region "Help Event Handler"
    ''' <summary>
    ''' Gives the information about the application
    ''' </summary>
    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        MessageBox.Show("This application is a text editor." & vbCrLf & vbCrLf &
                        "You can create new text file as well as edit your existing text file" & vbCrLf & vbCrLf &
                        "Created by Lovish Grover")
    End Sub
#End Region

#Region "File Event Handler"
    ''' <summary>
    ''' For creating a new file in the text editor. Clear the screen and reset all the variables to default 
    ''' </summary>
    Private Sub mnuFileNew_Click(sender As Object, e As EventArgs) Handles mnuFileNew.Click
        'resets all the varibales
        tbEditor.Clear()
        isFileSave = False
        openFilePath = String.Empty
        saveFilePath = String.Empty


    End Sub

    ''' <summary>
    ''' To load a file to the text editor box. Using Read To End function to load the data of file
    ''' some stream objects to load the path of the file and to read that path.  
    ''' </summary>
    Private Sub mnuFileOpen_Click(sender As Object, e As EventArgs) Handles mnuFileOpen.Click

        Dim openFile As FileStream
        Dim FileReader As StreamReader

        If opdOpen.ShowDialog() = DialogResult.OK Then

            openFilePath = opdOpen.FileName
            isFileSave = True

            openFile = New FileStream(openFilePath, FileMode.Open,
                                       FileAccess.Read)
            FileReader = New StreamReader(openFile)

            tbEditor.Text = FileReader.ReadToEnd

            FileReader.Close()
            saveFilePath = openFilePath

        End If
    End Sub

    ''' <summary>
    ''' To save the file, if file path already exists overrites the file 
    ''' if not creates a new path like the save as.
    ''' </summary>
    Private Sub mnuFileSave_Click(sender As Object, e As EventArgs) Handles mnuFileSave.Click

        If Not saveFilePath = String.Empty Then 'if there is file path overrite the file content 

            Dim write As New StreamWriter(saveFilePath)
            write.Write(tbEditor.Text)
            write.Close()
        Else
            mnuFileSaveAs_Click(sender, e) ' if there is no file path calls the save as functionality.
        End If

    End Sub


    ''' <summary>
    ''' To create a new file with new or existing data. Creates a new file path for the file saving
    ''' object file stream used to access the file path 
    ''' </summary>
    Private Sub mnuFileSaveAs_Click(sender As Object, e As EventArgs) Handles mnuFileSaveAs.Click

        Dim saveFile As FileStream
        Dim fileWriter As StreamWriter
        sfdSave.Filter = "Text Documents |*.txt|Word Document|*.docx|All Files|*.*" 'sets the filter for the files to save in this format

        If sfdSave.ShowDialog() = DialogResult.OK Then

            saveFilePath = sfdSave.FileName
            isFileSave = True

            saveFile = New FileStream(saveFilePath, FileMode.Create,
                                       FileAccess.Write)
            fileWriter = New StreamWriter(saveFile)

            fileWriter.Write(tbEditor.Text)

            fileWriter.Close()

        End If
    End Sub

    ''' <summary>
    ''' To close the editor
    ''' </summary>
    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click

        Me.Close()

    End Sub
#End Region

#Region "Edit Event Handler"

    ''' <summary>
    ''' To create a copy of the data existing in the text editor.
    ''' Copies the selected text into the clipboard and then make a copy of that.
    ''' </summary>
    Private Sub mnuEditCopy_Click(sender As Object, e As EventArgs) Handles mnuEditCopy.Click

        Clipboard.Clear() 'to clear the clipboard
        If tbEditor.SelectionLength > 0 Then
            My.Computer.Clipboard.SetText(tbEditor.SelectedText)

        End If

    End Sub

    ''' <summary>
    ''' To remove the selected text from the text editor and puts that text into 
    ''' set text method
    ''' </summary>
    Private Sub mnuEditCut_Click(sender As Object, e As EventArgs) Handles mnuEditCut.Click

        My.Computer.Clipboard.SetText(tbEditor.SelectedText)
        tbEditor.SelectedText = ""

    End Sub

    ''' <summary>
    ''' Insert the data in the text editor. Take data from the clipboard and insert where selected
    ''' </summary>
    Private Sub mnuEditPaste_Click(sender As Object, e As EventArgs) Handles mnuEditPaste.Click

        If Clipboard.ContainsText Then
            tbEditor.SelectedText = My.Computer.Clipboard.GetText()
        End If

    End Sub
#End Region
    Private Sub tbEditor_TextChanged(sender As Object, e As EventArgs) Handles tbEditor.TextChanged

        isFileSave = True

    End Sub


End Class
