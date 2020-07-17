Attribute VB_Name = "Module1"
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const MAX_PATH = 260
 
Public Sub CompactJetDatabase(Location As String, _
           Optional BackupOriginal As Boolean = True)
 On Error GoTo CompactErr
 Dim strBackupFile As String
 Dim strTempFile As String

 'Periksa apakah database ada...
 If Len(Dir(Location)) Then
    'Jika diperlukan utk membackup, lakukan!
    If BackupOriginal = True Then
        strBackupFile = GetTemporaryPath & "backup.mdb"
        If Len(Dir(strBackupFile)) Then
           Kill strBackupFile
        FileCopy Location, strBackupFile
        End If
    End If
    'Buat nama file temporal (sementara)
    strTempFile = GetTemporaryPath & "temp.mdb"
    If Len(Dir(strTempFile)) Then Kill strTempFile

    'Lakukan compact database menggunakan DBEngine
    DBEngine.CompactDatabase Location, strTempFile

    'Untuk repair database, Anda menggunakan cara
    'berikut:
    'Sesuaikan kebutuhan lainnya di prosedur ini...
    'DBEngine.RepairDatabase "NamaDatabaseAnda.mdb"

    'Jika database Anda dipassword, gunakan cari
    'berikut:
    'DBEngine.CompactDatabase Location, strTempFile, ,
    ', ";pwd=passwordanda;"

    'Hapus file database yang asli
    Kill Location
    'Copy yang file sementara dan telah dicompact
    'menjadi file database yang asli kembali...
    FileCopy strTempFile, Location
    'Hapus file database temporal (sementara)
    Kill strTempFile
    MsgBox "Sukses meng-compact database!", _
            vbInformation, "Sukses"
 End If
 Exit Sub
CompactErr:   'Jika terjadi error, tampilkan pesan
              'kemungkinan berikut ini...
Select Case Err.Number
       Case 70  'Sedang digunakan
            MsgBox "Database sedang digunakan!" & _
            vbCrLf & _
                   "Tutup dulu file tersebut!", _
            vbCritical, _
                    "Sedang Digunakan"
       Case 75  'Path/file belum ada
            MsgBox "Database belum dipilih." & _
                    vbCrLf & _
                    "Pilih dulu databasenya!", _
                    vbCritical, _
                   "Database Belum Ada"
       Case 3031  'Diprotect password
            MsgBox "Database dipassword," & vbCrLf & _
                    "lakukan langsung dari filenya!", _
                     vbCritical, _
                    "File Terprotect Password"
       Case 3343  'Database tidak dikenali
            MsgBox "Databaes bukan Access 97" & _
                     vbCrLf & _
                    "atau file bukan database!", _
                     vbCritical, _
                    "Database Tidak Dikenali"
       Case Else
            MsgBox Err.Number & " - " & Err.Description
     Exit Sub
  End Select

End Sub

'Fungsi ini untuk mengambil nama direktori tempat file
'database temporal (sementara) dicopy...
Public Function GetTemporaryPath()
  Dim strFolder As String
  Dim lngResult As Long
  strFolder = String(MAX_PATH, 0)
  lngResult = GetTempPath(MAX_PATH, strFolder)
  If lngResult <> 0 Then
    GetTemporaryPath = Left(strFolder, _
    InStr(strFolder, Chr(0)) - 1)
  Else
    GetTemporaryPath = ""
  End If
End Function


