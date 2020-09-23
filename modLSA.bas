Attribute VB_Name = "modLSA"
Option Explicit

'
' --- LSA I/O functions ---
'
' This module contains 2 functions (SetLsaData and GetLsaData) that can save/read unicode string to/from LSA.
' What is LSA? MS's definition of it, is
'
' ************************
' Local Security Authority
' (LSA) A protected subsystem that authenticates and logs users onto the local system.
' LSA also maintains information about all aspects of local security on a system, collectively
' known as the Local Security Policy of the system. LSA is available on operating systems in
' the Microsoft Windows .NET Server 2003 family, and on the Windows Advanced Server,
' Windows XP, Windows 2000, and Windows NT operating systems.
' ***********************************************************
'
' In LSA you can store passwords or whatever you want :-)
'
' written by Libor Blaheta
' credists : some functions in this module were written by Tadahiro Higuchi
'            his LSA sample codes can be found here - http://www1.harenet.ne.jp/cgi-bin/cgiwrap/unaap/chtml2/chtml.cgi?c1=3&c2=6&key=sec_lsa1
'

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function LsaOpenPolicy Lib "advapi32.dll" (ByVal pSystemName As Long, ByRef ObjectAttributes As LSA_OBJECT_ATTRIBUTES, ByVal DesiredAccess As Long, ByRef PolicyHandle As Long) As Long
Public Declare Sub LsaClose Lib "advapi32.dll" (ByVal hObjectHandle As Long)

Public Declare Function LsaRetrievePrivateData Lib "advapi32.dll" (ByVal PolicyHandle As Long, ByRef KeyName As LSA_UNICODE_STRING, ByRef pPrivateData As Long) As Long
Private Declare Function LsaStorePrivateData Lib "advapi32.dll" (ByVal PolicyHandle As Long, ByRef KeyName As LSA_UNICODE_STRING, ByRef PrivateData As LSA_UNICODE_STRING) As Long

Public Declare Function LsaFreeMemory Lib "advapi32.dll" (ByVal pBuffer As Long) As Long

Public Const ERROR_SUCCESS As Long = 0&
Public Const MAX_COMPUTERNAME As Long = 15
Public Const MAX_USERNAME As Long = 256
Public Const FILTER_NORMAL_ACCOUNT  As Long = &H2
Public Const MAX_PREFERRED_LENGTH As Long = -1

Public Type LSA_OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Public Type LSA_UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
    Buffer As String
End Type

Public Type LSA_UNICODE_STRING_LONG
    Length As Integer
    MaximumLength As Integer
    Buffer As Long
End Type

Public Const STATUS_SUCCESS As Long = &H0

Public Const READ_CONTROL As Long = &H20000
Public Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Public Const POLICY_VIEW_AUDIT_INFORMATION As Long = &H2&
Public Const POLICY_GET_PRIVATE_INFORMATION As Long = &H4&
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)
Public Const POLICY_TRUST_ADMIN As Long = &H8&
Public Const POLICY_CREATE_ACCOUNT As Long = &H10&
Public Const POLICY_CREATE_SECRET As Long = &H20&
Public Const POLICY_CREATE_PRIVILEGE As Long = &H40&
Public Const POLICY_SET_DEFAULT_QUOTA_LIMITS As Long = &H80&
Public Const POLICY_SET_AUDIT_REQUIREMENTS As Long = &H100&
Public Const POLICY_AUDIT_LOG_ADMIN As Long = &H200&
Public Const POLICY_SERVER_ADMIN As Long = &H400&

Public Const POLICY_WRITE As Long = (STANDARD_RIGHTS_WRITE Or POLICY_TRUST_ADMIN Or POLICY_CREATE_ACCOUNT Or POLICY_CREATE_SECRET Or POLICY_CREATE_PRIVILEGE Or POLICY_SET_DEFAULT_QUOTA_LIMITS Or POLICY_SET_AUDIT_REQUIREMENTS Or POLICY_AUDIT_LOG_ADMIN Or POLICY_SERVER_ADMIN)
Public Const POLICY_READ As Long = (STANDARD_RIGHTS_READ Or POLICY_VIEW_AUDIT_INFORMATION Or POLICY_GET_PRIVATE_INFORMATION)

'get data from the unicode structure

Public Function GetLsaStringLong(ByRef tLsaStringLong As LSA_UNICODE_STRING_LONG) As Byte()
Dim bData() As Byte
    
    With tLsaStringLong
        
        'check parameters
        If .Length <= 0 Or .Buffer <= 0 Then Exit Function
        
        'copy data
        ReDim bData(.Length - 1)
        CopyMemory ByVal VarPtr(bData(0)), ByVal .Buffer, .Length
        GetLsaStringLong = bData
    
    End With
    
End Function

'this functions makes a unicode string
Public Function MakeLsaString(ByVal sString As String) As LSA_UNICODE_STRING
    
    With MakeLsaString
        
        'empty string
        If sString = "" Then
            
            .Buffer = ""
            .Length = 0
            .MaximumLength = 0
        
        Else
            
            'convert string to unicode
            .Buffer = StrConv(sString, vbUnicode)
            'set lengths
            .Length = Len(sString) * 2
            .MaximumLength = (Len(sString) + 1) * 2
        
        End If
        
    End With
    
End Function

'store string in LSA
' hPolicy    - access rights
' sKeyName   - key name
' tLsaString - unicode string to save

Public Function SetLsaData(ByVal hPolicy As Long, ByVal sKeyName As String, tLsaString As LSA_UNICODE_STRING) As Boolean
Dim tLsaObjAttrib As LSA_OBJECT_ATTRIBUTES, hLSA As Long
Dim tLsaKey As LSA_UNICODE_STRING, lRet

    'open LSA
    If LsaOpenPolicy(0, tLsaObjAttrib, hPolicy, hLSA) <> STATUS_SUCCESS Then Exit Function
    
    'alse key name must be unicode string
    tLsaKey = MakeLsaString(sKeyName)
    
    'store string in LSA
    If LsaStorePrivateData(hLSA, tLsaKey, tLsaString) = STATUS_SUCCESS Then
        SetLsaData = True
    Else
        SetLsaData = False
    End If
    
    'close handle
    LsaClose hLSA
    
End Function

'read data from LSA
' hPolicy  - access rigths
' sKeyName - key name
'return value - byte array

Public Function GetLsaData(ByVal hPolicy As Long, ByVal sKeyName As String) As Byte()
Dim tLsaObjAttrib As LSA_OBJECT_ATTRIBUTES
Dim hLSA As Long, tLsaStringU As LSA_UNICODE_STRING
Dim pLsaData As Long, tLsaStringL As LSA_UNICODE_STRING_LONG

    'open LSA
    If LsaOpenPolicy(ByVal 0, tLsaObjAttrib, hPolicy, hLSA) <> STATUS_SUCCESS Then Exit Function

    'key name must be unicode string
    tLsaStringU = MakeLsaString(sKeyName)
        
    'read data
    If LsaRetrievePrivateData(hLSA, tLsaStringU, pLsaData) = STATUS_SUCCESS Then
    
        'copy structure
        CopyMemory ByVal VarPtr(tLsaStringL), ByVal pLsaData, Len(tLsaStringL)
        'get data from the structure
        GetLsaData = GetLsaStringLong(tLsaStringL)
        'free alocated memory
        LsaFreeMemory tLsaStringL.Buffer
    
    End If
    
    'free resources
    LsaFreeMemory VarPtr(tLsaStringU.Buffer)
    LsaClose hLSA

End Function

'get BSTR from pointer
'parameter - pData - pointer to BSTRstring
'return value - string
Public Function GetBSTR(ByVal pData As Long) As String
Dim iChar As Integer
   
    Do
        'get char
        CopyMemory iChar, ByVal pData, 2
        If iChar = 0 Then Exit Do
        'append char to the string
        GetBSTR = GetBSTR & Chr$(iChar)
        pData = pData + 2
    Loop

End Function
