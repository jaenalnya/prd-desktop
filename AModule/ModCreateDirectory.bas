Attribute VB_Name = "ModCreateDirectory"
Option Explicit
Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public NamaAplikasi As String
Public PathFileServer As String
Public PathFileLokal As String


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

Public Function isAdaUpdate(strFileLokal, strFileServer As String) As Boolean
On Error Resume Next
    Dim FSO As FileSystemObject
    Dim verServer As String
    Dim verLokal As String
    
    Set FSO = New FileSystemObject
    'Cek Versi File di Server
    verServer = FSO.GetFileVersion(strFileServer)
    'Cek Versi File di Lokal
    verLokal = FSO.GetFileVersion(strFileLokal)
    
    'Compare
    If verServer > verLokal Then
        isAdaUpdate = True
        Exit Function
    Else
        isAdaUpdate = False
    End If
End Function
Public Function Update(strFileLokal, strFileServer As String)
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    FSO.CopyFile strFileServer, strFileLokal, True
End Function

Public Function UpdateGambar(strFileLokal, strFileServer As String)
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    FSO.CopyFile strFileServer, strFileLokal, True
End Function

Public Function UploadGambar(strFileLokal, strFileServer As String)
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    FSO.CopyFile strFileLokal, strFileServer, True
End Function

