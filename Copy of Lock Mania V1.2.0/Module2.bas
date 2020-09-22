Attribute VB_Name = "Module2"
Option Explicit
Public M As Byte
Public LogM As Byte
Public PicM As Byte
Public StopFor As Boolean
Public StopDO As Boolean
Dim TEMP1 As String
Dim TEMP2 As String
Dim cont As Long
Public Remain As Integer
Public q As Long
Public PsWrd As String
Public KeyError As String
Public LenKey As Integer
Public tmplenkey As String
Public Rmvstr As String
Public i As Long, ii As Long, iii As Long, iiii As Long
Public Pici As Long
Dim picii
Dim a123, p123
Public RmvVal As String
Public Const TMPFILE As String = "C:\Not Of Ur Buisnes.TMP"
Public Const BKP As String = "C:\PIC.Bak"
Public Ext As String
Dim k
Dim Mi
Dim TmpStr As String
Public Report As String
Public picReport As String
Public TotalFileLen As Single
Public TotalWrittenBytes As Single
Dim per As String
Dim totper As String
Public aa As Long
Public bb As Long
Public c As Long
Public v As Long
Public stra As String
Public strb As String
Dim Dot As String
Dim PicExt As String
Dim PicExt2 As String
Dim ALP
Dim L As Integer, LL As Integer
Public TmpToFilePath As String
' #################### MSG COADER FUNCTIONS
Public Function ToHex(data As String, key As String) As String
key = CStr(GetKey(key))
For Mi = 1 To Len(data)
DoEvents
ToHex = ToHex & Hex(Asc(Mid(data, Mi, 1))) & CLng(key) & "_"
Next Mi
Write2Log (Now & vbCrLf & "LockMania Msg. Coder Used To Encrypt A Message")
End Function
Public Function ToStr(data As String, key As String) As String
Dim tmp As String
key = CStr(GetKey(key))

For Mi = 1 To Len(data)
On Error Resume Next
If Mid(data, Mi, 1) = "_" Then
ToStr = ToStr & Chr(Val("&H" + tmp))
tmp = ""
Else
tmp = tmp & Mid(data, Mi, 1)
If InStr(1, tmp, key) Then
tmp = Left(tmp, Len(tmp) - Len(key))
End If
End If
DoEvents
Next Mi
Write2Log (Now & vbCrLf & "LockMania Msg. Coder Used To Decrypt A Message")
End Function
Private Function GetKey(strkey As String) As Long
For k = 1 To Len(strkey)
DoEvents
GetKey = GetKey + Asc(Mid(strkey, k, 1))
Next k
End Function
'#####################  Files Code Function
Public Function RecsNum(NUM As Long, k As String) As Long
a123 = NUM / Len(k)
'Print a123
If InStr(1, a123, ".") = "0" Then
'MsgBox "OK"
RecsNum = a123
Else
Do
 NUM = NUM + 1
 a123 = NUM / Len(k)
 If InStr(1, a123, ".") = "0" Then
 Remain = NUM - FileLen(FrmMain.TxtFile.Text)
 'MsgBox "remain=" & Remain
 RecsNum = 0
 Exit Function
 End If
Loop
End If
End Function
Sub DOOPERATION(NumOfBytes As Long, password As String, LOCKFILE As Boolean, FilePath As String, NewFilePath As String)
Open FilePath For Binary As #1
Open NewFilePath For Binary As #2
'optimaize data....
Select Case LOCKFILE
Case True
For i = 1 To Len(password)
M = Asc(Mid(password, i, 1)) + 150
Put 2, i, M
Next i
For i = 1 To 3
M = Asc(Mid(Ext, i, 1))
Put 2, LOF(2) + 1, M
Next i
cont = 1
Case False
cont = Len(password) + 3 + 1
End Select
'emptying strings
TEMP1 = "": TEMP2 = ""
'################## START CODING
For i = 1 To NumOfBytes

For ii = cont To cont + (Len(password) - 1)
Get 1, ii, M
TEMP1 = TEMP1 & Chr(M)
Next ii

'MsgBox "TEMP1=" & TEMP1

'Put 3, 1, TEMP1
If LOCKFILE = True Then
For iii = 1 To Len(password)
q = Mid(password, iii, 1)
TEMP2 = TEMP2 & Mid(TEMP1, q, 1)
Next iii
Else
For iii = 1 To Len(password)
q = InStr(1, password, CStr(iii))
TmpStr = Mid(TEMP1, q, 1)
TEMP2 = TEMP2 & TmpStr
Next iii
End If

'MsgBox "TEMP2=" & TEMP2


For iiii = 1 To Len(TEMP2)
p123 = Mid(TEMP2, iiii, 1)
M = Asc(p123)
Put 2, LOF(2) + 1, M
Next iiii


TEMP1 = ""
TEMP2 = ""
cont = cont + Len(password)
DoEvents
FrmMain.LblPer.Caption = Int((LOF(2) / LOF(1)) * 100)
'#STOP FOR
If StopFor = True Then
msg = MsgBox("Are YOu Sure That You Want TO Stop Operations ?", vbYesNo + vbExclamation, "STOP!!")
Select Case msg
Case vbYes
StopOperation
'Exit For
Exit Sub
Case vbNo
StopFor = False
End Select
End If
Next i
' ############ END
If LOCKFILE = True Then
M = Asc(Len(password))
Put 2, LOF(2) + 1, M
End If
Close #1
Close #2

End Sub
Public Function checkpassword(pass As String) As Boolean
'First test " numbers only "
For i = 1 To Len(pass)
Select Case Mid(pass, i, 1)
Case 1 To 9    'clean password
Case 0
KeyError = "Please the Key Must Not Contain Zero"
checkpassword = False
Exit Function
Case Else     'bad password
KeyError = "Please the Key Must Contain Numbers Only"
checkpassword = False
Exit Function
End Select
' second test " Numbers Must be smaller then password length
If Int(Mid(pass, i, 1)) > Len(pass) Then
KeyError = "The Key Shouldnt Contain Numbers Bigger than It's Lenght"
checkpassword = False
Exit Function
End If
' third test " No Reapets "
q = 0
For ii = 1 To Len(pass)
If Mid(pass, i, 1) = Mid(pass, ii, 1) Then q = q + 1
If q = 2 Then
KeyError = "The Key Shouldnt Contain Reapeted Numbers"
checkpassword = False
Exit Function
End If
Next ii
Next i
checkpassword = True
End Function
Public Sub StopOperation()
Close #1
Close #2
On Error Resume Next
If Dir(TMPFILE) <> "" Then Kill TMPFILE
If Dir(FrmMain.TxtToFile.Text) <> "" Then Kill FrmMain.TxtToFile.Text
MsgBox "Operation Stopped", vbInformation, "Done"
FrmMain.CmdClear_Click
Write2Log (Now & vbCrLf & "Operation Stoped!!!")
End Sub
Public Sub ClearStrs()
TEMP1 = ""
TEMP2 = ""
M = 0
Remain = 0
q = 0
Rmvstr = ""
Ext = ""
Report = ""
End Sub
'################################ Picture Convertor Code
Public Function AddFile(clientfile As String, insertedfile As String) As String
PicExt = ""
PicExt2 = ""
On Error GoTo adderror:
c = "1"
Open clientfile For Binary As #20
aa = LOF(20)
Open insertedfile For Binary As #21
bb = LOF(21)
Do
Get 21, c, PicM
Put 20, aa + c, PicM
c = c + 1
per = CStr(Int(c / bb * 100))
totper = c + TotalWrittenBytes
totper = CStr(Int(totper / TotalFileLen * 100))
FrmMain.Lblpicper.Caption = per & " %"
FrmMain.lplpictotper.Caption = totper & " %"
'stop Do
If StopDO = True Then
msg = MsgBox("Are You Sure That U Want To Stop Operations ?", vbExclamation + vbYesNo, "Stop")
If msg = vbYes Then
StopPicOperations True
Exit Function
Else
StopDO = False
End If
End If
DoEvents
Loop Until c = bb + 1
Close #20
Close #21
Dot = InStrRev(insertedfile, ".", -1)
PicExt = Mid(insertedfile, Dot + 1, Len(insertedfile))

For picii = 1 To Len(PicExt)
ALP = Asc(Mid(PicExt, picii, 1))
ALP = ALP + 2
PicExt2 = PicExt2 & Chr(ALP)
Next picii

PicExt = PicExt2
TotalWrittenBytes = TotalWrittenBytes + bb
bb = aa + bb
'MsgBox aa
stra = aa / 2
strb = bb / 2
'MsgBox stra
FrmMain.LblPicStat.Caption = "WRITING TO LOG"
Write2Log Now & vbCrLf & "File : " & insertedfile & vbCrLf & "Hid To : " & clientfile
AddFile = PicExt & "(" & stra & "@" & strb & ")"
Exit Function
adderror:
msg = "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "WE ARE SO SORY But Try again"
MsgBox msg, vbCritical, "Error:"
End Function
Public Function extractfile(clientfile As String, extractedfile As String, password As String) As String
On Error GoTo extrerror:
PicExt = ""
PicExt2 = ""
L = InStr(1, password, "(")
PicExt = Left(password, L - 1)

For picii = 1 To Len(PicExt)
ALP = Asc(Mid(PicExt, picii, 1))
ALP = ALP - 2
PicExt2 = PicExt2 & Chr(ALP)
Next picii

PicExt = PicExt2

LL = InStr(1, password, "@")

stra = Mid(password, L + 1, LL - L - 1)
'MsgBox stra
stra = stra * 2
aa = CLng(stra)
strb = (Mid(password, LL + 1, Len(password) - LL - 1))
strb = strb * 2
bb = CLng(strb)
'MsgBox a & vbCrLf & b
c = "1"

Open clientfile For Binary As #22
Open extractedfile & "." & PicExt For Binary As #23

v = CLng(aa) + 1
Do
Get 22, v, PicM
Put 23, c, PicM
v = v + 1
c = c + 1
DoEvents
per = CStr(Int(v / bb * 100))
FrmMain.Lblpicper.Caption = per & " %"
If StopDO = True Then
msg = MsgBox("Are You Sure That U Want To Stop Operations ?", vbExclamation + vbYesNo, "Stop")
If msg = vbYes Then
StopPicOperations False
Exit Function
Else
StopDO = False
End If
End If
DoEvents
Loop Until v = bb + 1
Close #22
Close #23
FrmMain.LblPicStat.Caption = "WRITING TO LOG"
Write2Log (Now & vbCrLf & "File : " & extractedfile & "." & PicExt & vbCrLf & "Extracted FROM : " & clientfile)
extractfile = "File Extracted" & vbCrLf & "Path : " & extractedfile & "." & PicExt
Exit Function
extrerror:
msg = "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "WE ARE SO SORY But Try again"
MsgBox msg
End Function
Public Sub StopPicOperations(u As Boolean)
If u = True Then
Close #20
Close #21
On Error Resume Next
Kill FrmMain.TxtToFile.Text
If Dir(BKP) <> "" Then
FileCopy BKP, FrmMain.TxtPicFile.Text
Kill BKP
End If
Write2Log (Now & vbCrLf & "Picture Convertor Stoped While It Was Adding Files To A Picture")
FrmMain.Command8_Click
Else
On Error Resume Next
Kill TmpToFilePath
FrmMain.Command14_Click
Close #22
Close #23
Write2Log (Now & vbCrLf & "Picture Convertor Stoped While It Was Extracting Files From A Picture")
End If
End Sub

