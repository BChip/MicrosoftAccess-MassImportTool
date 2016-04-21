VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} automation_form 
   Caption         =   "Mass Importer/Exporter Tool - Made By: Bradley Chippi"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5670
   OleObjectBlob   =   "automation_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "automation_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'#########################################################################
'#AUTHOR: Bradley Chippi
'#PURPOSE: Automate Mass Imports and Exports into Access DB
'#########################################################################

'##############################HOW TO USE#################################
'#BE SURE TO ADD IN ALL THE REFERENCES
'#TO DO SO, FOLLOW THESE DIRECTIONS
'# 1) CLICK ON ON THE TOOLS TAB ON THE TOP
'# 2) CLICK ON REFERENCES...
'# 3) MAKE SURE THESE REFERENCES ARE CHECK MARKED AND ARE
'#    PLACED IN THIS ORDER:
'#      I)    Visual Basic For Applications
'#      II)   Microsoft Access 15.0 Object Library
'#      III)  Microsoft Office 15.0 Access database engine Object Library
'#      IV)   Microsoft ActiveX Data Objects 2.1 Library
'#      V)    Microsoft Visual Basic for Applications Extensibility 5.3
'#      VI)   OLE automation
'#      VII)  Microsoft Forms 2.0 Object Library
'#      VIII) Microsoft Office 15.0 Object Library
'#NOW CLICK ON THE GREEN PLAY BUTTON AT THE TOP
'#########################################################################

Option Compare Database
    Dim count As Integer

    Private Sub ToggleButton1_Click() 'WHEN START BUTTON IS CLICKED
    If OptionButton1 = True Then 'IF THE 'TEXT' RADIO BUTTON IS CHECKED
        If TextBox1.TextLength > 0 Then 'IF THERE IS TEXT IN TEXTBOX1
            TextFile 'JUMPTO TEXTFILE SUB
        Else
            MsgBox ("Where Do I Import From?")
        End If
    ElseIf OptionButton2 = True Then 'IF THE 'EXCEL' RADIO BUTTON IS CHECKED
        If TextBox1.TextLength > 0 Then 'IF THERE IS TEXT IN TEXTBOX1
            ExcelFile 'JUMPTO EXCELFILE SUB
        Else
            MsgBox ("Were Do I Import From?")
        End If
    Else
        MsgBox ("Please Pick Import Type") 'THIS DISPLAYS IF ONE RADIO BUTTON IS NOT CHECKED
    End If
    End Sub

    Private Sub ToggleButton2_Click() 'IF EXPORT BUTTON IS CLICKED
    If TextBox2.TextLength > 0 Then 'CHECKS IF TEXTBOX2 HAS TEXT
        Export
    Else
        MsgBox ("Where Do I Export To?")
    End If
End Sub

Private Sub ToggleButton3_Click() 'If first browse button is clicked
    GettingFolder1
End Sub

Private Sub ToggleButton4_Click()
    GettingFolder2
End Sub


Sub ExcelFile()
    Dim colFiles As New Collection 'MAKES A COLLECTION
    RecursiveDir colFiles, TextBox1.Text, "*.xlsx", True 'JUMP TO RECURSIVEDIR FUNCTION WITH PARAMETERS
    Dim WrdArray() As String 'MAKE A ARRAY
    Dim file() As String 'MAKE A ARRAY
    Dim vFile As Variant 'MAKE A VARIANT
    Dim cdb As DAO.Database
    Set cdb = CurrentDb 'DECLARE CDB IS OUR CURRENT DATABASE
    For Each vFile In colFiles 'FOR EACH FILE IN OUR FILE COLLECTION
        Me.Label2 = "Importing: " + vFile 'SET THE LABEL TO THE CURRENT FILE IN THE COLLECTION
        WrdArray() = Split(vFile, "\") 'SPLIT THE FILE PATH NAME BY "\" INTO A ARRAY
        file() = Split(WrdArray(count + 1), ".xlsx") 'SPLIT ARRAY AGAIN TO REMOVE .XLSX
        On Error Resume Next 'IF THERE IS A ERROR, PROCEED
        cdb.TableDefs.Delete file(0) 'DELETE TABLE BEING IMPORTED
        DoCmd.TransferSpreadsheet acImport, 9, file(0), vFile, True 'IMPORT
    Next vFile 'NEXT FILE IN COLLECTION
    Me.Label2 = "DONE!" 'DISPLAY DONE
End Sub

Sub TextFile() 'Same as ExcelFile
    Dim colFiles As New Collection
    RecursiveDir colFiles, TextBox1.Text, "*.txt", True
    Dim WrdArray() As String
    Dim file() As String
    Dim vFile As Variant
    Dim cdb As DAO.Database
    Set cdb = CurrentDb
    For Each vFile In colFiles
        Me.Label2 = "Importing: " + vFile
        WrdArray() = Split(vFile, "\")
        file() = Split(WrdArray(count + 1), ".txt")
        On Error Resume Next
        cdb.TableDefs.Delete file(0)
        DoCmd.TransferText acImportDelim, Null, file(0), vFile, True
    Next vFile
    Me.Label2 = "DONE!"
End Sub

Sub Export()
    Dim tbl As TableDef 'DECLARES TBL AS TABLEDEF
    For Each tbl In CurrentDb.TableDefs 'FOR EACH TABLE IN CURRENT DATABASE
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, tbl.Name, _
        TextBox2.Text + "\data.xls", True, tbl.Name 'EXPORT DATA TO EXCEL WORKBOOK AND EACH TABLE IS A EXCEL SHEET
    Next 'NEXT TABLE
    MsgBox ("DONE!")
End Sub

Sub GettingFolder1()
    Dim SelectedFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker) 'Open file dialog
    .Title = "Select folder" 'Sets title of the dialog
    .ButtonName = "Confirm" 'Sets button text
    .InitialFileName = "C:\" 'Sets starting location

    If .Show = -1 Then 'ok clicked
        SelectedFolder = .SelectedItems(1)
        count = Len(SelectedFolder) - Len(Replace(SelectedFolder, "\", "")) 'Count how many back-slashes - This makes program bullet proof
        TextBox1.Text = SelectedFolder
    Else 'cancel clicked
    End If

    End With

End Sub

Sub GettingFolder2() 'If second browser button is clicked
    Dim SelectedFolder As String

    With Application.FileDialog(msoFileDialogFolderPicker) 'Open file dialog
    .Title = "Select folder"
    .ButtonName = "Confirm"
    .InitialFileName = "C:\"

    If .Show = -1 Then 'ok clicked
    SelectedFolder = .SelectedItems(1)
    TextBox2.Text = SelectedFolder
    Else 'cancel clicked
    End If

    End With

End Sub


Public Function RecursiveDir(colFiles As Collection, _
                             strFolder As String, _
                             strFileSpec As String, _
                             bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add files in strFolder matching strFileSpec to colFiles
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders
        For Each vFolderName In colFolders
            Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
        Next vFolderName
    End If

End Function

Public Function TrailingSlash(strFolder As String) As String
    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = "\" Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & "\"
        End If
    End If
End Function

