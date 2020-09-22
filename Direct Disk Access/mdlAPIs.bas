Attribute VB_Name = "mdlAPIs"
Option Explicit

Public Const IOCTL_DISK_GET_DRIVE_GEOMETRY As Long = &H70000
Public Const IOCTL_STORAGE_MEDIA_REMOVAL As Long = &H2D4804

Public Const MAX_PATH As Long = 260
Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const OPEN_EXISTING As Long = 3
Public Const GENERIC_READ As Long = &H80000000
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const ERROR_FILE_NOT_FOUND As Long = 2


Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_CDROM = 5


Private PreventMediaRemoval As Byte

Enum MEDIATYPE
    unknown = 0
    F5_1Pt2_512 = 1
    F3_1Pt44_512 = 2
    F3_2Pt88_512 = 3
    F3_20Pt8_512 = 4
    F3_720_512 = 5
    F5_360_512 = 6
    F5_320_512 = 7
    F5_320_1024 = 8
    F5_180_512 = 9
    F5_160_512 = 10
    Removable = 11
    FixedMedia = 12
End Enum

Public Type DISK_GEOMETRY
   Cylinders         As Currency  'LARGE_INTEGER (8 bytes)
   MEDIATYPE         As Long
   TracksPerCylinder As Long
   SectorsPerTrack   As Long
   BytesPerSector    As Long
End Type

    
Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long

Public Declare Function GetLogicalDriveStrings Lib "kernel32" _
   Alias "GetLogicalDriveStringsA" _
  (ByVal nBufferLength As Long, _
   ByVal lpBuffer As String) As Long
  
Public Declare Function GetDriveType Lib "kernel32" _
   Alias "GetDriveTypeA" _
  (ByVal lpRootPathName As String) As Long

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



Public Sub GetDisksAndProfiles()
    'This function will Search for each Disk and then Call for DiskGeometry
    Dim hDevice As Long, devCount As Long
    
    Do
    
        'get handle of the Device. this is done by calling  CreateFile Function
        hDevice = CreateFile("\\.\PHYSICALDRIVE" & devCount, 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
        
        'Check if it's a valid handle
        If hDevice = INVALID_HANDLE_VALUE Then
            'We are Looking for more than one drives, so may be, this drive do not exists
            If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
                'There is No More Drive.
                Exit Do
            Else
                'Yes, there must be some Error
                MsgBox "Some Error Occured while attempting to open " & hDevice & " Disk." & vbCrLf & Err.Description, vbCritical
                Exit Do
            End If
        Else
            'Well, We have opened the Drive. Now Process it.
            Dim diskGeom As DISK_GEOMETRY
            diskGeom = GetDiskGeometry(hDevice)
            
            'Close the handle to the device.
            CloseHandle hDevice
            
            'Now Place the Values in the Text Box
            PrintValues diskGeom, devCount
            devCount = devCount + 1 'Proceed to NExt Disk
        End If
    Loop
    
End Sub


Public Function GetDiskGeometry(ByVal hDevice As Long) As DISK_GEOMETRY
    Dim totalBytes As Long, retVal As Long
    Dim diskGeom As DISK_GEOMETRY
    'Get the Geometry of the Disk
    retVal = DeviceIoControl(hDevice, IOCTL_DISK_GET_DRIVE_GEOMETRY, 0&, 0&, diskGeom, Len(diskGeom), totalBytes, 0&)
    
    GetDiskGeometry = diskGeom
End Function


Public Sub PrintValues(DG As DISK_GEOMETRY, ByVal devNo As Long)
    
    With frmMain.txtDiskG
        .Text = .Text & "Drive " & devNo & vbCrLf
        .Text = .Text & "Drive Type  : " & GetTypeOfDrive(DG.MEDIATYPE) & vbCrLf
        
        .Text = .Text & "Total Cylinders are    : " & CStr(DG.Cylinders * 10000) & vbCrLf
        .Text = .Text & "Tracks per Cyliner are : " & DG.TracksPerCylinder & vbCrLf
        .Text = .Text & "Sectors per Trak are   : " & DG.SectorsPerTrack & vbCrLf
        .Text = .Text & "Bytes Per Sector are   : " & DG.BytesPerSector & vbCrLf
        
        
    End With
End Sub

Public Function GetTypeOfDrive(ByVal mType As Long) As String
    
    Select Case mType
        Case MEDIATYPE.unknown
            GetTypeOfDrive = "Not Known"
        Case MEDIATYPE.F3_1Pt44_512
            GetTypeOfDrive = "3.5 inches, 1.44MB 512 bytes/sector"
        Case MEDIATYPE.F3_2Pt88_512
            GetTypeOfDrive = "3.5 inches, 2.88MB 512 bytes/sector"
        Case MEDIATYPE.F3_720_512
            GetTypeOfDrive = "3.5 inches, 720KB 512 bytes/sector"
            
        Case MEDIATYPE.F5_160_512
            GetTypeOfDrive = "5.25 inches, 160KB 512 bytes/sector"
        Case MEDIATYPE.F5_180_512
            GetTypeOfDrive = "5.25 inches, 180KB 512 bytes/sector"
        Case MEDIATYPE.F5_1Pt2_512
            GetTypeOfDrive = "5.25 inches, 1.2MB 512 bytes/sector"
        Case MEDIATYPE.F5_320_1024
            GetTypeOfDrive = "5.25 inches, 320KB 1024 bytes/sector"
        Case MEDIATYPE.F5_320_512
            GetTypeOfDrive = "5.25 inches, 320KB 512 bytes/sector"
        Case MEDIATYPE.F5_360_512
            GetTypeOfDrive = "5.25 inches, 360KB 512 bytes/sector"
        Case MEDIATYPE.FixedMedia
            GetTypeOfDrive = "Fixed Drive (HDD)"
        Case MEDIATYPE.Removable
            GetTypeOfDrive = "Removeable Drive"
    End Select
End Function


Public Function GetRemoveableDrives() As String()
    Dim sDriveBuffer As String
    Dim temp As Long, driveLetter As String
    Dim retVal As Long
    
    sDriveBuffer = Space(26 * 4)
    
    temp = 0
    
    'GetLogicalDriveStrings will Return Something like this A:\nullB:\nullc:\nullnull
    retVal = GetLogicalDriveStrings(Len(sDriveBuffer), sDriveBuffer)
    sDriveBuffer = Trim(sDriveBuffer)
    
    If retVal <> 0 Then 'If there is some drive/drives then
        Do Until driveLetter = Chr(0)
        
            driveLetter = Mid(sDriveBuffer, temp + 1, InStr(temp + 1, sDriveBuffer, Chr(0)) - 1)
            
            'Get the Location where Null is located
            temp = InStr(temp + 1, sDriveBuffer, Chr(0))
            
            'Get the Type of Drive, if that is Removable or CDROM then Add to List
            Dim dType As Long
            dType = GetDriveType(Left(driveLetter, 2))
            If dType = DRIVE_REMOVABLE Or dType = DRIVE_CDROM Then _
                frmMain.lstDrives.AddItem Left(driveLetter, 2)
        Loop
    End If
End Function


Public Function LockDrive(ByVal sDrive As String, ByVal bLock As Boolean) As Boolean
    Dim hDevice As Long, totalBytes As Long
    Dim retVal As Long
    
    hDevice = CreateFile("\\.\" & sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    PreventMediaRemoval = bLock 'CByte(Abs(bLock))
    retVal = DeviceIoControl(hDevice, IOCTL_STORAGE_MEDIA_REMOVAL, PreventMediaRemoval, Len(PreventMediaRemoval), 0&, 0&, totalBytes, 0&)
    
    CloseHandle hDevice
    
    LockDrive = CBool(retVal)
End Function

