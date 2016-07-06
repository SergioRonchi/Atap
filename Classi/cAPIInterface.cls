VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAPIInterface"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    LParam As Long
    iImage As Long
End Type



Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000




' Memory constants used through various memory API calls.
Private Const GMEM_MOVEABLE = &H2
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_ALL = &H10000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_WRITE = &H40000000

' The file/security API call constants.
' Refer to the MSDN for more information on how/what these constants
' are used for.
Private Const DACL_SECURITY_INFORMATION = &H4
Private Const SECURITY_DESCRIPTOR_REVISION = 1
Private Const SECURITY_DESCRIPTOR_MIN_LENGTH = 20
Private Const SD_SIZE = (65536 + SECURITY_DESCRIPTOR_MIN_LENGTH)
Private Const ACL_REVISION2 = 2
Private Const ACL_REVISION = 2
Private Const MAXDWORD = &HFFFFFFFF
Private Const SidTypeUser = 1
Private Const AclSizeInformation = 2

' The following are the inherit flags that go into the AceFlags field
' of an Ace header.

Private Const OBJECT_INHERIT_ACE = &H1
Private Const CONTAINER_INHERIT_ACE = &H2
Private Const NO_PROPAGATE_INHERIT_ACE = &H4
Private Const INHERIT_ONLY_ACE = &H8
Private Const INHERITED_ACE = &H10
Private Const VALID_INHERIT_FLAGS = &H1F
Private Const Delete = &H10000

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
'Display String resource ID or text in a pop-up window.
Private Const HH_DISPLAY_TEXT_POPUP = &HE
'Display mapped numeric value In dwData.
Private Const HH_HELP_CONTEXT = &HF
'Text pop-up help, similar To WinHelp's HELP_CONTEXTMENU.
Private Const HH_TP_HELP_CONTEXTMENU = &H10
'Text pop-up help, similar To WinHelp 's HELP_WM_HELP.
Private Const HH_TP_HELP_WM_HELP = &H11


Private Const IDH_Alias = 1
' Structures used by our API calls.
' Refer to the MSDN for more information on how/what these
' structures are used for.
Private Type ACE_HEADER
   AceType As Byte
   AceFlags As Byte
   AceSize As Integer
End Type


Private Type ACCESS_DENIED_ACE
  header As ACE_HEADER
  Mask As Long
  SidStart As Long
End Type

Private Type ACCESS_ALLOWED_ACE
   header As ACE_HEADER
   Mask As Long
   SidStart As Long
End Type

Private Type ACL
   AclRevision As Byte
   Sbz1 As Byte
   AclSize As Integer
   AceCount As Integer
   Sbz2 As Integer
End Type

Private Type ACL_SIZE_INFORMATION
   AceCount As Long
   AclBytesInUse As Long
   AclBytesFree As Long
End Type

Private Type SECURITY_DESCRIPTOR
   Revision As Byte
   Sbz1 As Byte
   Control As Long
   Owner As Long
   Group As Long
   sACL As ACL
   Dacl As ACL
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
' API calls used within this sample. Refer to the MSDN for more
' information on how/what these APIs do.

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (lpSystemName As String, ByVal lpAccountName As String, sID As Any, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Private Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As Byte, lpbDaclPresent As Long, pDacl As Long, lpbDaclDefaulted As Long) As Long
Private Declare Function GetFileSecurityN Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, ByVal pSecurityDescriptor As Long, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Long, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Long) As Long
Private Declare Function EqualSid Lib "advapi32.dll" (pSid1 As Byte, ByVal pSid2 As Long) As Long
Private Declare Function GetLengthSid Lib "advapi32.dll" (pSid As Any) As Long
Private Declare Function InitializeAcl Lib "advapi32.dll" (pAcl As Byte, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Private Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceIndex As Long, pace As Any) As Long
Private Declare Function AddAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, ByVal pAceList As Long, ByVal nAceListLength As Long) As Long
Private Declare Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As Byte, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Byte) As Long
Private Declare Function AddAccessDeniedAce Lib "advapi32.dll" (pAcl As Byte, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Byte) As Long
Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As Byte, ByVal bDaclDefaulted As Long) As Long
Private Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)



Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetVolumeInformation Lib "kernel32" _
    Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, _
    ByVal lpVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpfileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDiskFreeSpace _
    Lib "kernel32" _
    Alias "GetDiskFreeSpaceA" _
    (ByVal lpRootPathName As String, _
    lpSectorsPerCluster As Long, _
    lpBytesPerSector As Long, _
    lpNumberOfFreeClusters As Long, _
    lpTotalNumberOfClusters As Long) As Long
Private Declare Function HtmlHelp Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
    ByVal pszFile As String, ByVal uCommand As Long, _
    ByVal dwData As Long) As Long
Private Const S_OK = &H0                ' Success
Private Const S_FALSE = &H1             ' The Folder is valid, but does not exist
Private Const E_INVALIDARG = &H80070057 ' Invalid CSIDL Value


Private Const CSIDL_LOCAL_APPDATA = &H1C&
Private Const CSIDL_COMMON_APPDATA = &H23&
Private Const CSIDL_SYSTEM = &H25
Private Const CSIDL_FLAG_CREATE = &H8000&

Private Const SHGFP_TYPE_CURRENT = 0
Private Const SHGFP_TYPE_DEFAULT = 1
Private Const MAX_PATH = 260

Private Declare Function SHGetFolderPath Lib "shfolder" _
    Alias "SHGetFolderPathA" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, _
    ByVal hToken As Long, ByVal dwFlags As Long, _
    ByVal pszPath As String) As Long
Public Function GetSystemFolder() As String
Dim sPath As String
Dim RetVal As Long

' Fill our string buffer
sPath = String(MAX_PATH, 0)

RetVal = SHGetFolderPath(0, CSIDL_SYSTEM Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, sPath)

Select Case RetVal
    Case S_OK
        ' We retrieved the folder successfully
        
        ' All C strings are null terminated
        ' So we need to return the string upto the first null character
        sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
        GetSystemFolder = sPath
    Case S_FALSE
        ' The CSIDL in nFolder is valid, but the folder does not exist.
        ' Use CSIDL_FLAG_CREATE to have it created automatically
        err.Raise 1, , "The folder does not exist"
    Case E_INVALIDARG
        ' nFolder is invalid
        err.Raise 2, , "An invalid folder ID was specified"
    
End Select

End Function
Public Function GetCommonUserFolder() As String
Dim sPath As String
Dim RetVal As Long

' Fill our string buffer
sPath = String(MAX_PATH, 0)

RetVal = SHGetFolderPath(0, CSIDL_COMMON_APPDATA Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, sPath)

Select Case RetVal
    Case S_OK
        ' We retrieved the folder successfully
        
        ' All C strings are null terminated
        ' So we need to return the string upto the first null character
        sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
        GetCommonUserFolder = sPath
    Case S_FALSE
        ' The CSIDL in nFolder is valid, but the folder does not exist.
        ' Use CSIDL_FLAG_CREATE to have it created automatically
        err.Raise 1, , "The folder does not exist"
    Case E_INVALIDARG
        ' nFolder is invalid
        err.Raise 2, , "An invalid folder ID was specified"
    
End Select

End Function

Public Function GetCurrentUserFolder() As String
Dim sPath As String
Dim RetVal As Long

' Fill our string buffer
sPath = String(MAX_PATH, 0)

RetVal = SHGetFolderPath(0, CSIDL_LOCAL_APPDATA Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, sPath)

Select Case RetVal
    Case S_OK
        ' We retrieved the folder successfully
        
        ' All C strings are null terminated
        ' So we need to return the string upto the first null character
        sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
        GetCurrentUserFolder = sPath
    Case S_FALSE
        ' The CSIDL in nFolder is valid, but the folder does not exist.
        ' Use CSIDL_FLAG_CREATE to have it created automatically
        err.Raise 1, , "The folder does not exist"
    Case E_INVALIDARG
        ' nFolder is invalid
        err.Raise 2, , "An invalid folder ID was specified"
    
End Select
End Function

'Opens the compiled help file
Public Sub ShowHelpFile(strFilename As String, id As Long)
    Dim hwndHelp As Long
    Dim hwnd As Long
    'The return value is the window handle of the
    'created help window.
    hwndHelp = HtmlHelp(hwnd, strFilename, HH_HELP_CONTEXT, id)
End Sub
''
'EXposed constants
Public Property Get K_GENERIC_ALL() As String
 K_GENERIC_ALL = GENERIC_ALL
End Property

'''
''Exposed functions
Public Function EndCurrentProcess()
 ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
End Function
Public Function MySleep(milliseconds As Long)
 Sleep (milliseconds)
End Function

Public Function MyTextOut(ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
 MyTextOut = TextOut(hdc, x, Y, lpString, nCount)
End Function

Public Function Memory() As String
    Dim MemStat As MEMORYSTATUS
 
    GlobalMemoryStatus MemStat
    
    Memory = "Mem: " & Str$(Int(MemStat.dwTotalPhys / 1024 / 1024)) + " Mb (" + Str$(Int(MemStat.dwAvailPhys / 1024 / 1024)) + " Mb free);   Page File: " & Str$(Int(MemStat.dwTotalPageFile / 1024 / 1024)) + " Mb (" + Str$(Int(MemStat.dwAvailPageFile / 1024 / 1024)) + " MB free)"
End Function

Public Sub SetAccess(sUserName As String, sFileName As String, lMask As Long)
On Error GoTo fine
   Dim lResult As Long            ' Result of various API calls.
   Dim I As Integer               ' Used in looping.
   Dim bUserSid(255) As Byte      ' This will contain your SID.
   Dim bTempSid(255) As Byte      ' This will contain the Sid of each ACE in the ACL .
   Dim sSystemName As String      ' Name of this computer system.

   Dim lSystemNameLength As Long  ' Length of string that contains
                                  ' the name of this system.

   Dim lLengthUserName As Long    ' Max length of user name.

   'Dim sUserName As String * 255 ' String to hold the current user
                                  ' name.


   Dim lUserSID As Long           ' Used to hold the SID of the
                                  ' current user.

   Dim lTempSid As Long            ' Used to hold the SID of each ACE in the ACL
   Dim lUserSIDSize As Long          ' Size of the SID.
   Dim sDomainName As String * 255   ' Domain the user belongs to.
   Dim lDomainNameLength As Long     ' Length of domain name needed.

   Dim lSIDType As Long              ' The type of SID info we are
                                     ' getting back.

   Dim sFileSD As SECURITY_DESCRIPTOR   ' SD of the file we want.

   Dim bSDBuf() As Byte           ' Buffer that holds the security
                                  ' descriptor for this file.

   Dim lFileSDSize As Long           ' Size of the File SD.
   Dim lSizeNeeded As Long           ' Size needed for SD for file.


   Dim sNewSD As SECURITY_DESCRIPTOR ' New security descriptor.

   Dim sACL As ACL                   ' Used in grabbing the DACL from
                                     ' the File SD.

   Dim lDaclPresent As Long          ' Used in grabbing the DACL from
                                     ' the File SD.

   Dim lDaclDefaulted As Long        ' Used in grabbing the DACL from
                                     ' the File SD.

   Dim sACLInfo As ACL_SIZE_INFORMATION  ' Used in grabbing the ACL
                                         ' from the File SD.

   Dim lACLSize As Long           ' Size of the ACL structure used
                                  ' to get the ACL from the File SD.

   Dim pAcl As Long               ' Current ACL for this file.
   Dim lNewACLSize As Long        ' Size of new ACL to create.
   Dim bNewACL() As Byte          ' Buffer to hold new ACL.

   Dim sCurrentACE As ACCESS_ALLOWED_ACE    ' Current ACE.
   Dim pCurrentAce As Long                  ' Our current ACE.

   Dim nRecordNumber As Long

   ' Get the SID of the user. (Refer to the MSDN for more information on SIDs
   ' and their function/purpose in the operating system.) Get the SID of this
   ' user by using the LookupAccountName API. In order to use the SID
   ' of the current user account, call the LookupAccountName API
   ' twice. The first time is to get the required sizes of the SID
   ' and the DomainName string. The second call is to actually get
   ' the desired information.

   lResult = LookupAccountName(vbNullString, sUserName, _
      bUserSid(0), 255, sDomainName, lDomainNameLength, _
      lSIDType)

   ' Now set the sDomainName string buffer to its proper size before
   ' calling the API again.
   sDomainName = Space(lDomainNameLength)

   ' Call the LookupAccountName again to get the actual SID for user.
   lResult = LookupAccountName(vbNullString, sUserName, _
      bUserSid(0), 255, sDomainName, lDomainNameLength, _
      lSIDType)

   ' Return value of zero means the call to LookupAccountName failed;
   ' test for this before you continue.
     If (lResult = 0) Then
'        MsgBox "Error: Unable to Lookup the Current User Account: " _
'           & sUserName
        Exit Sub
     End If

   ' You now have the SID for the user who is logged on.
   ' The SID is of interest since it will get the security descriptor
   ' for the file that the user is interested in.
   ' The GetFileSecurity API will retrieve the Security Descriptor
   ' for the file. However, you must call this API twice: once to get
   ' the proper size for the Security Descriptor and once to get the
   ' actual Security Descriptor information.

   lResult = GetFileSecurityN(sFileName, DACL_SECURITY_INFORMATION, _
      0, 0, lSizeNeeded)

   ' Redimension the Security Descriptor buffer to the proper size.
   ReDim bSDBuf(lSizeNeeded)

   ' Now get the actual Security Descriptor for the file.
   lResult = GetFileSecurity(sFileName, DACL_SECURITY_INFORMATION, _
      bSDBuf(0), lSizeNeeded, lSizeNeeded)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
'      MsgBox "Error: Unable to Get the File Security Descriptor"
      Exit Sub
   End If

   ' Call InitializeSecurityDescriptor to build a new SD for the
   ' file.
   lResult = InitializeSecurityDescriptor(sNewSD, _
      SECURITY_DESCRIPTOR_REVISION)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
'      MsgBox "Error: Unable to Initialize New Security Descriptor"
      Exit Sub
   End If

   ' You now have the file's SD and a new Security Descriptor
   ' that will replace the current one. Next, pull the DACL from
   ' the SD. To do so, call the GetSecurityDescriptorDacl API
   ' function.

   lResult = GetSecurityDescriptorDacl(bSDBuf(0), lDaclPresent, _
      pAcl, lDaclDefaulted)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
'      MsgBox "Error: Unable to Get DACL from File Security " _
'         & "Descriptor"
      Exit Sub
   End If

   ' You have the file's SD, and want to now pull the ACL from the
   ' SD. To do so, call the GetACLInformation API function.
   ' See if ACL exists for this file before getting the ACL
   ' information.
   If (lDaclPresent = False) Then
'      MsgBox "Error: No ACL Information Available for this File"
      Exit Sub
   End If

   ' Attempt to get the ACL from the file's Security Descriptor.
   lResult = GetAclInformation(pAcl, sACLInfo, Len(sACLInfo), 2&)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
'      MsgBox "Error: Unable to Get ACL from File Security Descriptor"
      Exit Sub
   End If

   ' Now that you have the ACL information, compute the new ACL size
   ' requirements.
   lNewACLSize = sACLInfo.AclBytesInUse + (Len(sCurrentACE) + _
      GetLengthSid(bUserSid(0))) * 2 - 4

   ' Resize our new ACL buffer to its proper size.
   ReDim bNewACL(lNewACLSize)

   ' Use the InitializeAcl API function call to initialize the new
   ' ACL.
   lResult = InitializeAcl(bNewACL(0), lNewACLSize, ACL_REVISION)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
'      MsgBox "Error: Unable to Initialize New ACL"
      Exit Sub
   End If

   ' If a DACL is present, copy it to a new DACL.
   If (lDaclPresent) Then

      ' Copy the ACEs from the file to the new ACL.
      If (sACLInfo.AceCount > 0) Then

         ' Grab each ACE and stuff them into the new ACL.
         nRecordNumber = 0
         For I = 0 To (sACLInfo.AceCount - 1)

            ' Attempt to grab the next ACE.
            lResult = GetAce(pAcl, I, pCurrentAce)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
'               MsgBox "Error: Unable to Obtain ACE (" & I & ")"
               Exit Sub
            End If

            ' You have a pointer to the ACE. Place it
            ' into a structure, so you can get at its size.
            CopyMemory sCurrentACE, pCurrentAce, LenB(sCurrentACE)

            'Skip adding the ACE to the ACL if this is same usersid
            lTempSid = pCurrentAce + 8
            If EqualSid(bUserSid(0), lTempSid) = 0 Then

                ' Now that you have the ACE, add it to the new ACL.
                lResult = AddAce(VarPtr(bNewACL(0)), ACL_REVISION, _
                  MAXDWORD, pCurrentAce, _
                  sCurrentACE.header.AceSize)

                 ' Make sure you have the current ACE under question.
                 If (lResult = 0) Then
'                   MsgBox "Error: Unable to Add ACE to New ACL"
                    Exit Sub
                 End If
                 nRecordNumber = nRecordNumber + 1
            End If

         Next I

         ' You have now rebuilt a new ACL and want to add it to
         ' the newly created DACL.
         lResult = AddAccessAllowedAce(bNewACL(0), ACL_REVISION, _
            lMask, bUserSid(0))

         ' Make sure added the ACL to the DACL.
         If (lResult = 0) Then
'            MsgBox "Error: Unable to Add ACL to DACL"
            Exit Sub
         End If

         'If it's directory, we need to add inheritance staff.
         If GetAttr(sFileName) And vbDirectory Then

            ' Attempt to grab the next ACE which is what we just added.
            lResult = GetAce(VarPtr(bNewACL(0)), nRecordNumber, pCurrentAce)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
'               MsgBox "Error: Unable to Obtain ACE (" & I & ")"
               Exit Sub
            End If
            ' You have a pointer to the ACE. Place it
            ' into a structure, so you can get at its size.
            CopyMemory sCurrentACE, pCurrentAce, LenB(sCurrentACE)
            sCurrentACE.header.AceFlags = OBJECT_INHERIT_ACE + INHERIT_ONLY_ACE
            CopyMemory ByVal pCurrentAce, VarPtr(sCurrentACE), LenB(sCurrentACE)

            'add another ACE for files
            lResult = AddAccessAllowedAce(bNewACL(0), ACL_REVISION, _
               lMask, bUserSid(0))

            ' Make sure added the ACL to the DACL.
            If (lResult = 0) Then
'               MsgBox "Error: Unable to Add ACL to DACL"
               Exit Sub
            End If

            ' Attempt to grab the next ACE.
            lResult = GetAce(VarPtr(bNewACL(0)), nRecordNumber + 1, pCurrentAce)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
'               MsgBox "Error: Unable to Obtain ACE (" & I & ")"
               Exit Sub
            End If

            CopyMemory sCurrentACE, pCurrentAce, LenB(sCurrentACE)
            sCurrentACE.header.AceFlags = CONTAINER_INHERIT_ACE
            CopyMemory ByVal pCurrentAce, VarPtr(sCurrentACE), LenB(sCurrentACE)
        End If


         ' Set the file's Security Descriptor to the new DACL.
         lResult = SetSecurityDescriptorDacl(sNewSD, 1, _
            bNewACL(0), 0)

         ' Make sure you set the SD to the new DACL.
         If (lResult = 0) Then
'            MsgBox "Error: " & _
'                "Unable to Set New DACL to Security Descriptor"
            Exit Sub
         End If

         ' The final step is to add the Security Descriptor back to
         ' the file!
         lResult = SetFileSecurity(sFileName, _
            DACL_SECURITY_INFORMATION, sNewSD)

         ' Make sure you added the Security Descriptor to the file!
         If (lResult = 0) Then
'            MsgBox "Error: Unable to Set New Security Descriptor " _
'               & " to File : " & sFileName
'            MsgBox err.LastDllError
         Else
'            MsgBox "Updated Security Descriptor on File: " _
'               & sFileName
         End If

      End If

   End If
fine:
End Sub





Public Function GetFormatData() As String
    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim Buffer As String, sT As SYSTEMTIME
    With sT
        .wDay = 6
        .wMonth = 8
        .wYear = 2000
    End With
    Buffer = String(255, 0)
    GetDateFormat ByVal 0&, 0, sT, vbNullString, Buffer, Len(Buffer)
    Buffer = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
    GetFormatData = Buffer
End Function

    
    
Public Function GetDriveInfo(filePath As String) As String
    Dim lpVolumeNameBuffer As String
    Dim nVolumeNameSize As String
    Dim lpVolumeSerialNumber As Long
    Dim lpMaximumComponentLength As Long
    Dim lpfileSystemFlags As Long
    Dim lpFileSystemNameBuffer As String
    Dim nFileSystemNameSize As Long
    Dim lpSectorsPerCluster As Long
    Dim lpBytesPerSector As Long
    Dim lpNumberOfFreeClusters As Long
    Dim lpTotalNumberOfClusters As Long
    Dim rtval As Long
    lpVolumeNameBuffer = Space(255)
    nVolumeNameSize = 256
    lpFileSystemNameBuffer = Space(255)
    nFileSystemNameSize = 256
    rtval = GetVolumeInformation(filePath, _
        lpVolumeNameBuffer, _
        nVolumeNameSize, _
        lpVolumeSerialNumber, _
        lpMaximumComponentLength, _
        lpfileSystemFlags, _
        lpFileSystemNameBuffer, _
        nFileSystemNameSize)
    rtval = GetDiskFreeSpace(filePath, _
        lpSectorsPerCluster, _
        lpBytesPerSector, _
        lpNumberOfFreeClusters, _
        lpTotalNumberOfClusters)
    
    GetDriveInfo = "Volume Serial Number:        " & lpVolumeSerialNumber & _
                    Chr(13) & Chr(10) & _
                    "Volume Size In KBytes:        " & lpTotalNumberOfClusters / 1024 * lpSectorsPerCluster * lpBytesPerSector & _
                    " Kbytes" & Chr(13) & Chr(10) & _
                    "Volume Free Space In KBytes:  " & lpNumberOfFreeClusters / 1024 * lpSectorsPerCluster * lpBytesPerSector / 1024 & " Kbytes"
    
End Function
Public Function BrowseFolder(ByVal hwnd As Long, ByVal szDialogTitle As String) As String
    Dim Esito As Long
    Dim Bi As BROWSEINFO
    Dim dwIList As Long
    Dim szPath As String
    
    Bi.hOwner = hwnd
    Bi.lpszTitle = szDialogTitle
    Bi.ulFlags = BIF_RETURNONLYFSDIRS
    szPath = Space$(1024)
    dwIList = SHBrowseForFolder(Bi)
    Esito = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If Esito Then
        BrowseFolder = Left$(szPath, InStr(szPath, Chr(0)) - 1)
    Else
        BrowseFolder = ""
    End If

End Function
Public Function GetTimeStamp(Optional ExcludeData As Boolean) As String
  
  Dim s As SYSTEMTIME
  GetSystemTime s

If ExcludeData Then
  GetTimeStamp = Format(Now, "HHMMSS") & String(3 - Len(CStr(s.wMilliseconds)), "0") & s.wMilliseconds
Else
   GetTimeStamp = Format(Now, "YYYYMMDDHHMMSS") & String(3 - Len(CStr(s.wMilliseconds)), "0") & s.wMilliseconds
End If
End Function
cessHandle = 0
End Function



