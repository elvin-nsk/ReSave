VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "ReSaveCdr"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Dim FileName As String
Dim MouseX As Single, MouseY As Single

'===============================================================================

Private Sub UserForm_Initialize()
    VersionsList.Clear
    Dim i As Long
    Dim LowestVersion As Long
    If Application.VersionMajor > 21 Then
        LowestVersion = 15
    Else
        LowestVersion = 13
    End If
    For i = LowestVersion To Application.VersionMajor - 1
      VersionsList.AddItem "Version " & i
    Next i
    VersionsList.ListIndex = GetSetting("ReSaveCDR", "Setting", "version", "0")
    If VersionsList.ListIndex = -1 Then VersionsList.ListIndex = 0
    
    FromTextBox.Text = GetSetting("ReSaveCDR", "Setting", "from", "")
    ToTextBox.Text = GetSetting("ReSaveCDR", "Setting", "to", "")
    
    lblBarFront.Width = 0
End Sub

Private Sub UserForm_Terminate()
    Call SaveSetting("ReSaveCDR", "Setting", "version", VersionsList.ListIndex)
    Call SaveSetting("ReSaveCDR", "Setting", "from", FromTextBox.Text)
    Call SaveSetting("ReSaveCDR", "Setting", "to", ToTextBox.Text)
End Sub

Private Sub VersionsList_Click()
    Dim i As Integer
    If VersionsList.Value = True Then
        For i = 0 To FilesList.ListCount - 1
        FilesList.Selected(i) = True
        Next i
    Else
      For i = 0 To FilesList.ListCount - 1
        FilesList.Selected(i) = False
      Next i
    End If
      
    NumberOfFilesLabel.Caption = "Files " & FilesList.ListCount
End Sub

Private Sub CmxCheckBox_Click()
    If CmxCheckBox.Value = True Then
        WmfCheckBox.Value = False: EmfCheckBox.Value = False: VersionsList.Enabled = False
    End If
    If CmxCheckBox.Value = False And WmfCheckBox.Value = False And EmfCheckBox.Value = False Then
        VersionsList.Enabled = True
    End If
End Sub

Private Sub WmfCheckBox_Click()
    If WmfCheckBox.Value = True Then
        EmfCheckBox.Value = False: CmxCheckBox.Value = False: VersionsList.Enabled = False
    End If
    If CmxCheckBox.Value = False And WmfCheckBox.Value = False And EmfCheckBox.Value = False Then
        VersionsList.Enabled = True
    End If
End Sub

Private Sub EmfCheckBox_Click()
    If EmfCheckBox.Value = True Then
        WmfCheckBox.Value = False: CmxCheckBox.Value = False: VersionsList.Enabled = False
    End If
    If CmxCheckBox.Value = False And WmfCheckBox.Value = False And EmfCheckBox.Value = False Then
        VersionsList.Enabled = True
    End If
End Sub

Private Sub BrowseFromFolder_Click()
    On Error Resume Next
    Dim Str As String
    Str = CorelScriptTools.GetFolder(FromTextBox.Text, "Select a folder")
    FromTextBox.Text = Str
    FileName = VBA.Dir(FromTextBox.Text & "\" & "*.CDR")
    FilesList.Clear
       
    Do While FileName <> ""
        If FileName <> "." And FileName <> ".." Then
            If (VBA.GetAttr(FromTextBox.Text & FileName) And vbDirectory) = vbDirectory Then
                FilesList.AddItem FileName
                If SelecatAllCheckBox.Value Then
                    FilesList.Selected(FilesList.ListCount - 1) = True
                End If
            End If
        End If
        FileName = VBA.Dir
    Loop
       
    Call VersionsList_Click
End Sub

Private Sub BrowseToFolder_Click()
    Dim Str As String
    Str = CorelScriptTools.GetFolder(ToTextBox.Text, "Select a folder for resave")
    ToTextBox.Text = Str
End Sub

Private Sub ReSaveButton_Click()
    If FromTextBox.Text = "" Then MsgBox "Select a folder!", vbExclamation, "Me CDR": Exit Sub
    If ToTextBox.Text = "" Then MsgBox "Select a folder!", vbExclamation, "Me CDR": Exit Sub
    
    On Error GoTo ErrHandler:
    
    Optimization = True
    Dim Opt As New StructSaveAsOptions, doc As Document, ExpOpt As StructExportOptions
    Dim ExpFlt As ExportFilter
    Set ExpOpt = CreateStructExportOptions
    ExpOpt.UseColorProfile = False
    
     Opt.EmbedICCProfile = False
     Opt.EmbedVBAProject = False
     Opt.Filter = cdrCDR
     Opt.IncludeCMXData = False
     Opt.Overwrite = True
     Opt.Range = cdrAllPages
     Opt.ThumbnailSize = cdr10KColorThumbnail
     Opt.Version = VBA.Split(VersionsList.Text, " ")(1)
     Dim xxx As Double
     Dim i As Integer, jj As Integer
     Dim Ret As VbMsgBoxResult
     
     If FilesList.ListCount = 0 Then Exit Sub
     xxx = 216 / (FilesList.ListCount)
     
     jj = 0
     Dim d As Document
     Set d = Nothing
     Dim newstr As String
     
       For i = 0 To FilesList.ListCount - 1
             If FilesList.Selected(i) = True Then
                 Set d = Nothing
                 Set d = OpenDocument(FromTextBox.Text & "\" & FilesList.List(i))
                 Dim nn As Integer
                 nn = d.ActivePage.Shapes.Count
                     If CmxCheckBox.Value = False And WmfCheckBox.Value = False And EmfCheckBox.Value = False Then
                         d.SaveAs ToTextBox.Text & "\" & FilesList.List(i), Opt
                     ElseIf CmxCheckBox.Value = True Then 'cmx
                        newstr = Replace(ToTextBox.Text & "\" & FilesList.List(i), ".cdr", ".cmx")
                        Set ExpFlt = ActiveDocument.ExportEx(newstr, cdrCMX6, cdrAllPages, ExpOpt)
                        ExpFlt.Finish
                     ElseIf WmfCheckBox.Value = True Then 'wmf
                        newstr = Replace(ToTextBox.Text & "\" & FilesList.List(i), ".cdr", ".wmf")
                        Set ExpFlt = ActiveDocument.ExportEx(newstr, cdrWMF, cdrAllPages, ExpOpt)
                        ExpFlt.Finish
                     ElseIf EmfCheckBox.Value = True Then 'emf
                        newstr = Replace(ToTextBox.Text & "\" & FilesList.List(i), ".cdr", ".emf")
                        Set ExpFlt = ActiveDocument.ExportEx(newstr, cdrEMF, cdrAllPages, ExpOpt)
                        ExpFlt.Finish
                    End If
                     d.Close
                  End If
                 lblBarFront.Width = i * xxx
                 DoEvents
NextFile1:
       Next i
       
        lblBarFront.Width = 0
        Optimization = False
        If jj > 0 Then MsgBox "Unsaved files: " & jj
        Application.Refresh
        Exit Sub
    
ErrHandler:
        jj = jj + 1
        Resume NextFile1

End Sub

Private Sub FilesList_Change()
    Dim m As Integer, mm As Integer
    mm = 0
    For m = 0 To FilesList.ListCount - 1
        If FilesList.Selected(m) = True Then
            mm = mm + 1
        End If
        
    Next m
    NumberOfSelectedLabel.Caption = "Selected: " & mm
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        Me.Left = Me.Left - MouseX + X
        Me.Top = Me.Top - MouseY + Y
    End If
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

Private Sub CorelVba_Click()
    With VBA.CreateObject("WScript.Shell")
        .Run "http://corelvba.com/"
    End With
End Sub

Private Sub ElvinNsk_Click()
    With VBA.CreateObject("WScript.Shell")
        .Run "https://vk.com/elvin_macro/ReSaveCdr"
    End With
End Sub

'===============================================================================

Private Sub FormCancel()
    Me.Hide
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(Ñancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Ñancel = True
        FormCancel
    End If
End Sub
