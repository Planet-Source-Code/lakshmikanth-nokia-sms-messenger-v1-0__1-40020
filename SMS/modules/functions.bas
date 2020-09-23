Attribute VB_Name = "functions"
'================================================================
'================================================================
'***                  SMS Messenger v1.0                      ***
'================================================================
'================================================================
'*** For any Questions or Comments concerning this program    ***
'*** Homepage : http://xt.tp//www.sms.com                     ***
'*** Email    : kanthji@rediffmail.com                        ***
'*** Yahoo    : lucky_kanth                                   ***
'================================================================
'================================================================
'*** Mobule Name : functions.bas                              ***
'*** Date & Time : 21st October 2002, 04:30(IST)              ***
'================================================================
'================================================================

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Ssend As Integer
Public Srow As Integer
Public id As Integer
Public tes As String

Public Log As String

Public Function sendsms(SMSC, Phno, SMS)

   Srow = Srow + 1
   SMSmainfrm.SMSgrid.TextMatrix(Srow, 0) = "Sending.."
   SMSmainfrm.SMSgrid.TextMatrix(Srow, 1) = Phno
   SMSmainfrm.SMSgrid.TextMatrix(Srow, 2) = SMS
   If SMSmainfrm.Option1.Value = True Then
      SMSmainfrm.SMSgrid.TextMatrix(Srow, 3) = "NORMAL"
   Else
      If SMSmainfrm.Option2.Value = True Then
         SMSmainfrm.SMSgrid.TextMatrix(Srow, 3) = "FLASH"
      Else
         If SMSmainfrm.Option3.Value = True Then
            SMSmainfrm.SMSgrid.TextMatrix(Srow, 3) = "BLINK"
         End If
      End If
   End If
   SMSmainfrm.SMSgrid.Rows = SMSmainfrm.SMSgrid.Rows + 1

  If SMSmainfrm.DReport.Value = 1 Then DR = "31" Else DR = "11"
  If Len(Phno) < 12 Then FO = "81" Else FO = "91"
  If SMSmainfrm.Option1.Value = True Then
     mf = "00"
  Else
    If SMSmainfrm.Option2.Value = True Then
       mf = "F0"
    Else
      If SMSmainfrm.Option3.Value = True Then
         mf = "08"
      End If
    End If
  End If
  SMSC = revnum(SMSC)
  Phno = revnum(Phno)
  If mf = "08" Then
     SMS = ConvToblinkhex(SMS)
     LS = hex2(Len(SMSmainfrm.Text1.Text) * 2)
  Else
     SMS = ConvToHex(SMS)
     LS = hex2(Len(SMSmainfrm.Text1.Text))
  End If
  mva = hex2(SMSmainfrm.Combo1.ListIndex)
  tep = DR & "00" & hex2(Len(Phno)) & FO & Phno & "00" & mf & mva & LS & SMS
  lnth = (Len(tep)) / 2
  msg = "0791" & SMSC & tep
  If Ssend = 1 Then
     Call sendmsg(lnth, msg)
     SMSmainfrm.SMSgrid.TextMatrix(Srow, 0) = "SENT"
  Else
     Call savemsg(lnth, msg)
     SMSmainfrm.SMSgrid.TextMatrix(Srow, 0) = "SAVED"
  End If
  'SMSmainfrm.Text1.Text = SMSmainfrm.SMS.Input

End Function
Public Function sendmsg(lnth, msg)
  SMSmainfrm.SMS.Output = "AT" & Chr$(13) & Chr(10)
  SMSmainfrm.SMS.Output = "AT+CMGF=0" & Chr$(13) & Chr(10)
  SMSmainfrm.SMS.Output = "AT+CMGS=" & lnth & Chr$(13) & Chr(10)
  SMSmainfrm.SMS.Output = msg & Chr$(26)
End Function

Public Function savemsg(lnth, msg)
  SMSmainfrm.SMS.Output = "AT" & Chr$(13) & Chr(10)
  SMSmainfrm.SMS.Output = "AT+CMGF=0" & Chr$(13) & Chr(10)
  SMSmainfrm.SMS.Output = "AT+CMGW=" & lnth & ",3" & Chr$(13) & Chr(10)
  SMSmainfrm.SMS.Output = msg & Chr$(26)
End Function
'reversing phone numbers
Public Function revnum(numb)
     s = 1
     ma = ""
     While (s <= Len(numb))
       ta = Mid(numb, s, 2)
       A = Mid(ta, 1, 1)
       B = Mid(ta, 2, 1)
       If B = "" Then B = "F"
       ma = ma & B & A
       s = s + 2
     Wend
     revnum = ma
End Function
Public Function ConvToblinkhex(Text)
    Dim MainTextTemp As String
    MainTextTemp = Text
    Text = ""
    
    For B = 1 To Len(MainTextTemp)
       If Mid(MainTextTemp, B, 1) = "|" Then
         X = "01"
       Else
         X = hex2(Asc(Mid(MainTextTemp, B, 1)))
       End If
        Text = Text & "00" & X
        X = ""
    Next B
    
    ConvToblinkhex = Text
End Function


Public Function ConvToHex(Text)
'This converts text to hex

    Dim B, g As Integer
    Dim MainTextTemp As String
    Dim X As String
    ReDim lbit(160), rbit(160), fin(140) As String
    MainTextTemp = Text
    
    g = 8
    For B = 1 To Len(MainTextTemp)
        If g = 0 Then
           g = 8
        End If
        X = binary(Asc(Mid(MainTextTemp, B, 1)))
        If g <= 8 Then
              lbit(B) = Mid(X, g, 7)
              rbit(B) = Mid(X, 1, g - 1)
              g = g - 1
        End If
        X = ""
    Next B
    J = 0
    For i = 0 To Len(MainTextTemp)
       If i <> Len(MainTextTemp) Then
          If (lbit(i + 1) & rbit(i)) <> "" Then
              fin(J) = lbit(i + 1) & rbit(i)
              J = J + 1
          End If
       Else
          If (rbit(i)) <> "" Then
             fin(J) = rbit(i)
          End If
       End If
    Next i
    Text = ""
    For i = 0 To J
       If fin(i) <> "" Then
          Text = Text & Hex1(fin(i))
       End If
    Next i
    ConvToHex = Text
    MainTextTemp = ""
End Function

Public Function findstr(strp, str)
    txt = strp
    B = 1
    For B = 1 To Len(txt)
       D = Mid(txt, B, Len(str))
       If D = str Then findstr = 1
    Next B
End Function

Public Function binary(Ascii As Integer)
Dim A As Integer, B As Integer, C As Integer, D As Integer, E As Integer, F As Integer, g As Integer, h As Integer
Dim Final As String
If Ascii < 128 Then
  A = 0
Else
  A = 1
  Ascii = Ascii - 128
End If
If Ascii < 64 Then
  B = 0
Else
  B = 1
  Ascii = Ascii - 64
End If
If Ascii < 32 Then
  C = 0
Else
  C = 1
  Ascii = Ascii - 32
End If
If Ascii < 16 Then
  D = 0
Else
  D = 1
  Ascii = Ascii - 16
End If
If Ascii < 8 Then
  E = 0
Else
  E = 1
  Ascii = Ascii - 8
End If
If Ascii < 4 Then
  F = 0
Else
  F = 1
  Ascii = Ascii - 4
End If
If Ascii < 2 Then
  g = 0
Else
  g = 1
  Ascii = Ascii - 2
End If
If Ascii < 1 Then
  h = 0
Else
  h = 1
End If
binary = B & C & D & E & F & g & h

End Function
Public Function Hex1(y)
  s = Len(y)
  r = ""
  A = 0
  B = 1
  While (s > 0)

     r = Mid(y, s, 1) * B
     s = s - 1
     B = B * 2
     A = A + r
  Wend
  
  Hex1 = hex2(A)
  
End Function
Public Function hex2(he)
    y = Hex(he)
    If Len(y) = 1 Then
       y = "0" & y
    End If
    hex2 = y
End Function


Public Function addcombo()
   h = 1
   For h = 1 To 3
       SMSmainfrm.Combo1.AddItem h
   Next h

   SMSmainfrm.Combo2.AddItem "9600,n,8,1"
   SMSmainfrm.Combo2.AddItem "19200,n,8,1"
   SMSmainfrm.Combo2.AddItem "38400,n,8,1"
   SMSmainfrm.Combo2.AddItem "57600,n,8,1"

   h = 0
   Hrs = 0
   Min = 1
   sp = 4
   For h = 0 To 143
       SMSmainfrm.Combo3.AddItem Space(sp) & h & " - " & Hrs & " hrs " & Min * 5 & " Min"
       Min = Min + 1
       If Min = 12 Then
          Min = 0
          Hrs = Hrs + 1
       End If
       If h = 9 Then sp = 2
       If h = 99 Then sp = 0
   Next h
   Hrs = 12
   Min = 30
   For h = 144 To 167
       SMSmainfrm.Combo3.AddItem h & " - " & Hrs & " hrs " & Min & " Min"
       If Min = 30 Then
          Min = 0
          Hrs = Hrs + 1
       Else
          Min = 30
          'Hrs = Hrs + 1
       End If
   Next h
   days = 2
   For h = 168 To 196
       SMSmainfrm.Combo3.AddItem h & " - " & days & " days"
       days = days + 1
   Next h
   weeks = 5
   For h = 197 To 255
       SMSmainfrm.Combo3.AddItem h & " - " & weeks & " weeks"
       weeks = weeks + 1
   Next h
  
   SMSmainfrm.Combo1.ListIndex = 2
   SMSmainfrm.Combo2.ListIndex = 0
   SMSmainfrm.Combo3.ListIndex = 167

End Function
Public Sub SaveValue(Key As String, Value As String)
    WritePrivateProfileString "SMS Messenger v1.0", Key, Value, App.Path & "\settings.ini"
End Sub

Sub GetPrefs()
   SMSmainfrm.Combo1.ListIndex = ReadValue("Port")
   SMSmainfrm.Combo2.ListIndex = ReadValue("Settings")
   SMSmainfrm.Combo3.ListIndex = ReadValue("Msgvalidity")
   SMSmainfrm.Text3.Text = ReadValue("SMSC")
End Sub


Sub SetPrefs()
   SaveValue "Port", SMSmainfrm.Combo1.ListIndex
   SaveValue "Settings", SMSmainfrm.Combo2.ListIndex
   SaveValue "Msgvalidity", SMSmainfrm.Combo3.ListIndex
   SaveValue "SMSC", SMSmainfrm.Text3.Text
End Sub

Public Function ReadValue(Key As String, Optional Default As String, Optional Section As String = "SMS Messenger v1.0", Optional File)
    Dim sReturn As String
    If IsMissing(File) Then File = FullPath(App.Path, "settings.ini")
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), File))
End Function

Function FullPath(lpPath As String, lpFile As String) As String
If Right(lpPath, 1) <> "\" Then lpPath = lpPath & "\"
FullPath = lpPath & lpFile
End Function
