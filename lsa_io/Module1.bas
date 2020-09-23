Attribute VB_Name = "mod"
Option Explicit

'
' --- LSA I/O sample ---
'
' This code saves a string in LSA and read it.
' For more informations look at the module modLSA.bas
'
' You should be logged as Admin otherwise you won't be able to access LSA.
'
' written by Libor Blaheta
'

'key name
Private Const KEY_NAME As String = "L$_some_key_name"

Private Sub Main()
Dim x As LSA_UNICODE_STRING, abData() As Byte, sOut As String, i As Long

    'make unicode string
    x = MakeLsaString("Test string")
    
    'write our data
    If SetLsaData(POLICY_WRITE, KEY_NAME, x) = True Then
    
        'read stored data
        abData = GetLsaData(POLICY_READ, KEY_NAME)
    
        'filet null chars
        For i = LBound(abData) To UBound(abData) Step 2
            sOut = sOut & Chr(abData(i))
        Next i
    
        'info
        MsgBox "Retrieved data - " & sOut, vbInformation, "LSA I/O"
        
    Else
    
        MsgBox "Data was not saved, maybe you are not logged as Admin.", vbInformation, "LSA I/O"
        
    End If
        
End Sub
