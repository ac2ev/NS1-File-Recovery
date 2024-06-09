Attribute VB_Name = "modINI"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)


Public Function GetIniParam(NomFichier As String, NomSection As String, NomVariable As String) As String
  Dim ReadString As String * 255
  Dim returnv    As String
  Dim mResultLen As Integer

  mResultLen = GetPrivateProfileString(NomSection, NomVariable, "(Unassigned)", ReadString, Len(ReadString) - 1, NomFichier)
  If IsNull(ReadString) Or Left(ReadString, 12) = "(Unassigned)" Then
     Dim Tempvalue As Variant
     Dim Message As String
     Message = "Le fichier de configutation " & NomFichier & " est introuvable."
     returnv = ""
  Else
     returnv = Left(ReadString, InStr(ReadString, Chr$(0)) - 1)
  End If
  GetIniParam = returnv
End Function

Public Function WriteWinIniParam(NomDuIni As String, sLaSection As String, sNouvelleCle As String, sNouvelleValeur As String)
Dim iSucccess As Integer
    
    iSucccess = WritePrivateProfileStringByKeyName(sLaSection, sNouvelleCle, sNouvelleValeur, NomDuIni)
    If iSucccess = 0 Then
        MsgBox "L'édition du fichier a échoué.", vbCritical, "Erreur"
        WriteWinIniParam = False
    Else
        WriteWinIniParam = True
    End If

End Function

Function Encrypte(sData As String) As String
    Dim sTemp As String, sTemp1 As String
    Dim iI%, lT

    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) + 10
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Encrypte = sTemp1$
End Function


Function Decrypt(sData As String) As String
    Dim sTemp As String, sTemp1 As String
    Dim iI%, lT

    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) - 10
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Decrypt = sTemp1$
End Function

Public Function RetShortName(vPath As String)
'Retourne la string droite d'une chaine

Dim intLenght As Integer
Dim intPos    As Integer
For intLenght = 1 To Len(vPath)
    If Left(Right(vPath, intLenght), 1) = "\" Then
       intPos = intLenght
       Exit For
    End If
Next intLenght
If intPos = 0 Then
   RetShortName = vPath
Else
   RetShortName = Right(vPath, intPos - 1)
End If
End Function

