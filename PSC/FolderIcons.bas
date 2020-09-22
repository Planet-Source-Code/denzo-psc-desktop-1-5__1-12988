Attribute VB_Name = "FolderIcons"
Public Sub SetFolderIcon(ByVal Folder As String, ByVal Icon As String, ByVal IconIndex As Integer)
'Dim Folder As String
'Folder = Foldertxt.Text
On Error Resume Next
Dim File_NUM As Integer, Line_Text As String, Text_Buff As String
File_NUM = FreeFile: Text_Buff = ""
'Exists...
 If StrConv(Dir(Folder & "\desktop.ini", vbSystem), vbUpperCase) <> "" Or _
    StrConv(Dir(Folder & "\desktop.ini", vbArchive), vbUpperCase) <> "" Or _
    StrConv(Dir(Folder & "\desktop.ini", vbHidden), vbUpperCase) <> "" Then
    Dim num_f As Integer
    num_f = FreeFile
    Open Folder & "\Desktop.ini" For Input As #num_f
    While Not EOF(num_f)
    Dim Pass_Icon As Boolean
    Pass_Icon = False
      'Replace the old lines...
          Line Input #num_f, Line_Text
          If Left(StrConv(Line_Text, vbUpperCase), 9) = "ICONFILE=" Then
              Text_Buff = Text_Buff & vbCrLf & "IconFile=" & Icon
              Pass_Icon = True
          ElseIf Left(StrConv(Line_Text, vbUpperCase), 8) = "ICOINDEX" Then
              Text_Buff = Text_Buff & vbCrLf & "IcoIndex=" + Format(IconIndex)
          Else
              If Line_Text <> vbCrLf And Len(Line_Text) > 1 Then _
                  Text_Buff = Text_Buff & vbCrLf & Line_Text
          End If
      Wend
      Close num_f
      If Not Pass_Icon Then _
        Text_Buff = "[.ShellClassInfo]" & vbCrLf & _
          "IconFile=" & Icon & vbCrLf & _
          "IcoIndex=" + Format(IconIndex)
    Else
        Text_Buff = "[.ShellClassInfo]" & vbCrLf & _
        "IconFile=" & Icon & vbCrLf & _
             "IcoIndex=" + Format(IconIndex)
    End If
Write_File Folder & "\Desktop.ini", Text_Buff
'Now we've to attrib the folder +s -> system :]
' without doing this the icon won't appear !!
Attribs Folder, "+s"
End Sub

Private Sub Write_File(MyFileToWrite As String, TextToWrite As String)
On Error Resume Next
Dim num_fi As Integer
num_fi = FreeFile
'Set normal attributes to the file...
    Attribs MyFileToWrite, "-s -h -r"
'Now let's write it. using the buffer... ;)
    Kill MyFileToWrite
    Open MyFileToWrite For Append As num_fi
        Print #num_fi, TextToWrite
    Close num_fi
End Sub
Private Sub Attribs(MyFileOrDir As String, MyOptions As String)
Dim File_NUM As Integer
File_NUM = FreeFile
Open "C:\r001.bat" For Output As #File_NUM
    Print #File_NUM, "@echo off" & vbCrLf & _
     "attrib " & MyOptions & " """ & _
     MyFileOrDir & """" & vbCrLf & "exit"
Close File_NUM
Shell "C:\r001.bat", vbMinimizedNoFocus
End Sub

