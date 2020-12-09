Attribute VB_Name = "Module1"

' WriteMess Function
' ==================
' Written by: Mike
' Function: Write the sent string to the log files
'
Public Sub writeup(mess As String)
Dim iFile As Integer
Dim sfile As String

sfile = App.Path & "\logs\pklog." & Format(Now(), "yyy-mm-dd") & ".txt."
iFile = FreeFile()
Open sfile For Append As #iFile
Print #iFile, Now() & " " & mess
Close #1

End Sub

