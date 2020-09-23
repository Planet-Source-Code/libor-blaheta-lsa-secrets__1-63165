VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RAS secrets in LSA"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      Height          =   885
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   7335
   End
   Begin VB.ListBox lstUsers 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SID && RAS secrest"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   7365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Users"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' --- RAS Secrets in LSA ---
'
' Windows NT allows users to save their RAS credentials by using the 'Save Password' checkbox when making a dial-up connection.
' Then credintials are stored in LSA under these keys "RasCredentials!SID#0", RasDialParams!SID#0 and L$_RasDefaultCredentials#0. Where "SID" is user's SID.
'
' This VB app can read values of these keys so you can see RAS logins and passwords.
'
' Some info about LSA and these keys can be found here - http://www.insecure.org/sploits/Win95.safe-password-fraud.html
'
' note : You should be logged as Admin otherwise you won't be able to access LSA.
'
' Libor Blaheta
'

Private Declare Function NetUserEnum Lib "NETAPI32.dll" (ByVal servername As String, ByVal level As Long, ByVal filter As Long, ByRef pBuffer As Long, ByVal prefmaxlen As Long, ByRef entriesread As Long, ByRef totalentries As Long, ByRef resume_handle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Long, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare Function ConvertSidToStringSidW Lib "advapi32.dll" (ByVal pSid As Long, ByRef pStringSid As Long) As Long
Private Declare Function IsValidSid Lib "advapi32.dll" (ByVal pSid As Long) As Long

Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Type USER_INFO
    sUserName As String   'user name
    sSID As String        'SID
End Type

Dim atUI() As USER_INFO

Private Sub Form_Load()
Dim asUsers() As String, i As Long

    'get all accounts
    Call EnumUsers

    'fill the listbox with accounts names
    For i = LBound(atUI) To UBound(atUI)
        'get account's sid
        atUI(i).sSID = GetLocalSid(atUI(i).sUserName)
        lstUsers.AddItem atUI(i).sUserName
    Next i
     
    'default credentials
    Label2.Caption = "L$_RasDefaultCredentials#0 = " & GetTextFromLsa("L$_RasDefaultCredentials#0")

End Sub

'
' this function returns SID of user-name
' parameter - sUserName - user name
' return value - SID
'

Private Function GetLocalSid(Optional sUserName As String = vbNullString) As String
Dim lUserNameSize As Long, lSidSize As Long, lDomainSize As Long
Dim Sid(0 To 255) As Byte
Dim snu As Integer, pSid As Long
Dim sDomain As String

    If sUserName = vbNullString Then
    
        'get user-name
        sUserName = Space(256)
        lUserNameSize = Len(sUserName)
        If GetUserName(sUserName, lUserNameSize) = 0 Then Exit Function
    
        sUserName = Left(sUserName, lUserNameSize - 1)
    
    End If
        
    lSidSize = 255
    sDomain = Space(255)
    lDomainSize = Len(sDomain)
    
    'get SID
    If LookupAccountName(vbNullString, sUserName, ByVal VarPtr(Sid(0)), lSidSize, sDomain, lDomainSize, snu) = 0 Then Exit Function
    'check it
    If IsValidSid(VarPtr(Sid(0))) = 0 Then Exit Function
    'convert it to string
    If ConvertSidToStringSidW(VarPtr(Sid(0)), pSid) = 0 Then Exit Function
    
    'get string from pointer to it
    GetLocalSid = GetBSTR(pSid)
    'free alocated memory
    LocalFree pSid
    
End Function

'
' this function enumerates all accounts on
' return value - no. accoutns and names are stored in the array atUI
'

Private Function EnumUsers() As Long
Dim pData As Long, pName As Long
Dim lEntries As Long, lTotalEntries As Long
Dim i As Long, sNames() As String

    'get accounts
    If NetUserEnum(vbNullString, 0, FILTER_NORMAL_ACCOUNT, pData, MAX_PREFERRED_LENGTH, lEntries, lTotalEntries, ByVal 0) = ERROR_SUCCESS Then
   
        If lEntries = 0 Then GoTo NOENTRY
        ReDim atUI(0 To lEntries - 1)
        
        For i = 0 To lEntries - 1
      
            'copy names to array
            CopyMemory ByVal VarPtr(pName), ByVal pData + i * Len(pName), Len(pName)
            atUI(i).sUserName = GetBSTR(pName)
        
        Next i
        
   End If
   
NOENTRY:

   EnumUsers = lEntries
   'free alocated memory
   NetApiBufferFree pData

End Function

Private Sub lstUsers_Click()
Dim abData() As Byte, i As Long
Dim sData As String, s As String
    
    If lstUsers.ListIndex <> -1 Then
        
Dim sCred As String, sParam As String
        
        'get data from LSA
        sParam = GetTextFromLsa("RasDialParams!" & atUI(lstUsers.ListIndex).sSID & "#0")
        sCred = GetTextFromLsa("RasCredentials!" & atUI(lstUsers.ListIndex).sSID & "#0")
        
        'check return value
        If sParam = "" Then sParam = "No RAS secrets"
        If sCred = "" Then sCred = "No RAS secrets"
        
        txtInfo.Text = "SID = " & atUI(lstUsers.ListIndex).sSID & vbCrLf & "RasDialParams = " & sParam & vbCrLf & "RasCredentials = " & sCred
        
    End If
        
End Sub

'
' this function reads data from LSA and convert them to VB string
' parameters - sKey - LSA key name
' return value - read data

Private Function GetTextFromLsa(ByVal sKey As String) As String
Dim abData() As Byte, i As Long
Dim s As String

    'get data
    abData = GetLsaData(POLICY_GET_PRIVATE_INFORMATION, sKey)
        
    'check the array
    On Error Resume Next
        i = LBound(abData)
        If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    'filter NULLS
    For i = LBound(abData) To UBound(abData)
        If abData(i) <> 0 Then s = s & Chr(abData(i))
    Next i
        
    GetTextFromLsa = s
        
End Function


