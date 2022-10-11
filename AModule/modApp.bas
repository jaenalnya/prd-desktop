Attribute VB_Name = "modApp"
Option Explicit
Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Const sAS_AutoBackup As String = "AutoBackup"

Public Function SetAutoBackup(ByVal NewValue As Boolean)

    Dim sValue As String
    
    SaveSetting App.Title, "AppSetting", sAS_AutoBackup, IIf(NewValue, "T", "F")

End Function

Function BackupDatabase(dbname1 As String, ipath As String, ByRef bckfile As String, Optional ByRef errmsg As String = "") As Boolean

On Error GoTo xc
Dim newbck
newbck = dbname1 & "_" & Format(Date, "mm-dd-yyyy") & "_" & Format(Now, "hhmmss") & ".bak"
bckfile = newbck
DoEvents
CN.Execute "BACKUP DATABASE [" & dbname1 & "] TO  DISK = N'" & ipath & newbck & "' WITH NOFORMAT, INIT,  NAME = N'" & dbname1 & "-Full Database Backup', SKIP, NOREWIND, NOUNLOAD"
BackupDatabase = True
Exit Function

xc:
BackupDatabase = False
End Function

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Function RestoreDatabase(bckfile As String, DBName As String, Optional ByRef errmsg As String) As Boolean
On Error GoTo xc
DoEvents
CN.Execute "RESTORE DATABASE " & DBName & " FROM DISK = '" & bckfile & "' WITH REPLACE"

RestoreDatabase = True

Exit Function

xc:
Debug.Print "RESTORE DATABASE " & DBName & " FROM DISK = '" & bckfile & "' WITH REPLACE"
errmsg = Err.Description
RestoreDatabase = False
End Function


'Fungsi mencek keberadaan folder
Public Function DirectoryExist(DirPath As String) As Boolean
DirectoryExist = Dir(DirPath, vbDirectory) <> ""
End Function

'Fungsi untuk membuat Folder
Public Sub CreateNewDirectory(NewDirectory As String)
Dim sDirTest As String
Dim SecAttrib As SECURITY_ATTRIBUTES
Dim bSuccess As Boolean
Dim sPath As String
Dim iCounter As Integer
Dim sTempDir As String

sPath = NewDirectory

If Right(sPath, Len(sPath)) <> "\" Then
sPath = sPath & "\"
End If

iCounter = 1

Do Until InStr(iCounter, sPath, "\") = 0
iCounter = InStr(iCounter, sPath, "\")
sTempDir = Left(sPath, iCounter)
sDirTest = Dir(sTempDir)
iCounter = iCounter + 1
'create directory
SecAttrib.lpSecurityDescriptor = &O0
SecAttrib.bInheritHandle = False
SecAttrib.nLength = Len(SecAttrib)
bSuccess = CreateDirectory(sTempDir, SecAttrib)
Loop
End Sub

'Fungsi Untuk Menghapus folder
Public Sub DelDirectory(sName As String)
On Error Resume Next
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
If Dir(sName, vbDirectory) <> "" Then
FSO.DeleteFolder sName
End If
Set FSO = Nothing
End Sub

