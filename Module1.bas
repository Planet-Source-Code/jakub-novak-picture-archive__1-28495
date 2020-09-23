Attribute VB_Name = "Module1"
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Const DRIVE_CDROM = 5
Public CDDrives() As tCDInfo
Public Type tCDInfo
        CDSerial As Long
        CDLabel As String
        CDDrive As String
End Type
Public Type tFileInfo
    fNAME As String
    fDATE As String
    fSize As String
    FullPath As String
End Type
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

Public gCancel As Boolean

Public Function GetDriveSerial(driveroot As String) As Long
Dim x As String
Dim Y As Long
Dim strVoleSerial As Long
    
GetVolumeInformation driveroot, x, 255, GetDriveSerial, 255, &H1, x, Y
    
End Function
Public Function GetDrivesInfo()
Dim tmp, Y, i
Dim sDrives As String
Dim x() As tCDInfo
Dim DrivesCount
Dim lpFSNB, nFSNS, lpVNB
Dim CDSerial
Dim CDDrive, CDLabel As String
    
    DrivesCount = 0
    sDrives = String(255, 0)
    GetLogicalDriveStrings 255, sDrives
    sDrives = Mid(sDrives, 1, InStrRev(sDrives, "\"))
    tmp = Split(sDrives, Chr(0))
    For i = 0 To UBound(tmp)
        
        If GetDriveType(tmp(i)) = DRIVE_CDROM Then
            DrivesCount = DrivesCount + 1
            ReDim Preserve x(0 To DrivesCount)
            CDDrive = tmp(i)
            CDLabel = String(255, " ")
            GetVolumeInformation CDDrive, CDLabel, 255, CDSerial, 255, &H1, lpFSNB, 255
            CDLabel = Replace(Trim(CDLabel), Chr(0), "")
            x(DrivesCount).CDDrive = CDDrive
            x(DrivesCount).CDSerial = CDSerial
            x(DrivesCount).CDLabel = CDLabel
        End If
    Next i
    'GetDriveListing = X
    If DrivesCount = 0 Then MsgBox ("No CD Drive found in Your PC.")
    ReDim CDDrives(DrivesCount)
    CDDrives = x
End Function
 Public Function GetVolumeLabel(driveroot) As String
 Dim x As String
    x = String(255, " ")
    Debug.Print GetVolumeInformation(driveroot, x, 255, vbNull, 255, &H1, sss, 255)
    GetVolumeLabel = Replace(Trim(x), Chr(0), "")
    
 End Function

Public Function FindFiles(sPath, Files() As tFileInfo)
Dim FoundFile As WIN32_FIND_DATA
Dim FileName As String
Dim Ext As String
Dim FileCount
Dim ffhwnd As Long

If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
ffhwnd = FindFirstFile(sPath & "*", FoundFile)
FileName = Trim(Replace(FoundFile.cFileName, Chr(0), ""))

FileCount = 0
Do While FileName <> ""
    Ext = LCase(Right(FileName, 3))
    If FileName = "." Or FileName = ".." Then GoTo NextFile
    If Ext <> "jpg" And Ext <> "bmp" And Ext <> "gif" And _
    GetFileAttributes(sPath & FileName) <> 16 And _
    GetFileAttributes(sPath & FileName) <> 17 Then GoTo NextFile
    
    If GetFileAttributes(sPath & FileName) = 16 Or GetFileAttributes(sPath & FileName) = 17 Then
        'Folder
        FileName = UCase(FileName)
       Else
        'File
        FileName = LCase(FileName)
    End If

    ReDim Preserve Files(0 To FileCount)
    Files(FileCount).fNAME = FileName
    Files(FileCount).FullPath = sPath & FileName
    If GetFileAttributes(sPath & FileName) <> 16 And _
       GetFileAttributes(sPath & FileName) <> 17 Then Files(FileCount).fSize = FileLen(sPath & FileName)
        
    FileCount = FileCount + 1
NextFile:
FoundFile.cFileName = Chr(0)
FindNextFile ffhwnd, FoundFile
FileName = Trim(Replace(FoundFile.cFileName, Chr(0), ""))
Loop
sPath = Mid(sPath, 1, Len(sPath) - 1)
End Function

Private Function CheckDirContent(sPath)
Dim x

If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

x = Dir(sPath & "*.*")

End Function

Public Function CreateDB(FileName)
Dim MyDB As Database
Dim MyTable As TableDef
Dim MyField As Field
Dim MyIndex As Index
Dim SQL
Dim MyQuery As QueryDef


On Error GoTo err1
Set MyDB = DBEngine.CreateDatabase(FileName, dbLangGeneral)

'Create table CDs
Set MyTable = MyDB.CreateTableDef("CDs")
Set MyField = MyTable.CreateField("CDID", dbLong)
MyField.Attributes = 17
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("CDLabel", dbText, 50)
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("CDSerial", dbDouble, 50)
MyTable.Fields.Append MyField
MyDB.TableDefs.Append MyTable

'Creat Primary index
Set MyIndex = MyTable.CreateIndex("primarykey")
Set MyField = MyIndex.CreateField("CDID")
MyIndex.Primary = True
MyIndex.Required = True
MyIndex.Fields.Append MyField
MyTable.Indexes.Append MyIndex


'Create table Files
Set MyTable = MyDB.CreateTableDef("Files")
Set MyField = MyTable.CreateField("ID", dbLong)
MyField.Attributes = 17
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("CDID", dbLong)
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("Name", dbText, 255)
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("FullPath", dbText, 255)
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("Type", dbByte)
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("Thumb", dbLongBinary)
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("Dimension", dbText, 20)
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("Description", dbText, 80)
MyField.AllowZeroLength = True
MyTable.Fields.Append MyField
Set MyField = MyTable.CreateField("KeyWords", dbMemo, 255)
MyField.AllowZeroLength = True
MyTable.Fields.Append MyField
MyDB.TableDefs.Append MyTable

'Creat Primary index
Set MyIndex = MyTable.CreateIndex("primarykey")
Set MyField = MyIndex.CreateField("ID")
MyIndex.Primary = True
MyIndex.Required = True
MyIndex.Fields.Append MyField
MyTable.Indexes.Append MyIndex




SQL = "SELECT Files.ID"
SQL = SQL & " FROM Files, Files AS Files_1"
SQL = SQL & " Where (((InStr([Files_1].[FullPath], [Files].[FullPath])) = 1) And ((Files_1.Type) = 0))"
SQL = SQL & " GROUP BY Files.ID"
SQL = SQL & " ORDER BY Files.ID;"
Set MyQuery = MyDB.CreateQueryDef("DirsToLeave", SQL)

SQL = "SELECT Files.ID"
SQL = SQL & " FROM Files LEFT JOIN DirsToLeave ON Files.ID = DirsToLeave.ID"
SQL = SQL & " WHERE (((DirsToLeave.ID) Is Null));"
Set MyQuery = MyDB.CreateQueryDef("DirsToDELETE", SQL)


Set MyIndex = Nothing
Set MyField = Nothing
Set MyTable = Nothing
MyDB.Close
Set MyDB = Nothing

CreateDB = True

Exit Function


err1:
CreateDB = False
MsgBox Error

End Function

