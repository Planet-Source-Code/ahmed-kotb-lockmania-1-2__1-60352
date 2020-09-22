Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public LogPath As String
Dim iset As Long
Public msg As String
Public song As String
Public Dlg1 As New dlg
Public logi As Long
Public Function PlayMidiFile(FilePath As String) As Boolean

    Dim iset As Long

    On Error Resume Next

    If Dir(FilePath) = "" Then Exit Function
    
    'Stop mid
    iset = mciSendString("stop midi", "", 0, 0)
    iset = mciSendString("close midi", "", 0, 0)

    'Play
    iset = mciSendString("open sequencer!" & FilePath & " alias midi", "", 0, 0)
    iset = mciSendString("play midi", "", 0, 0)
    PlayMidiFile = (iset = 0)

End Function
Public Sub StopMusic()
    iset = mciSendString("stop midi", "", 0, 0)
    iset = mciSendString("close midi", "", 0, 0)
End Sub
'############### LOG CODE
Public Sub Write2Log(str As String)
On Error Resume Next
Dim T As Long
Open LogPath For Binary As #44
str = vbCrLf & vbCrLf & str
T = LOF(44) + 1
For logi = T To T + Len(str) - 1
LogM = Asc(Mid(str, (logi - T) + 1, 1))
Put 44, logi, LogM
DoEvents
Next logi
Close #44
End Sub
Public Function loadLog() As String
loadLog = ""
Open LogPath For Binary As #44
For logi = 1 To LOF(44)
Get 44, logi, LogM
loadLog = loadLog & Chr(LogM)
Next logi
Close #44
End Function
Public Sub Clearlog()
Open LogPath For Output As #44
Close #44
End Sub
Public Sub GradientPlus(frm As PictureBox, StartRed As Integer, _
StartGreen As Integer, StartBlue As Integer, _
EndRed As Integer, EndGreen As Integer, EndBlue As Integer)

    On Error Resume Next
    Dim X As Integer
    Dim RedChange As Integer
    Dim GreenChange As Integer
    Dim BlueChange As Integer
    frm.DrawStyle = 6 ' Inside Solid
    frm.ScaleMode = 3 ' Pixels
    frm.DrawMode = 13 ' Copy Pen
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For X = 0 To 255 'Start Loop
    
        frm.Line (0, X)-(Screen.Width, X - 1), _
        RGB(StartRed + RedChange, StartGreen + GreenChange, _
        StartBlue + BlueChange), B 'Draws Line With correct color
        
        RedChange = RedChange + (EndRed - StartRed) / 255 '
        GreenChange = GreenChange + (EndGreen - StartGreen) / 255 ' Sets Next Loops Color
        BlueChange = BlueChange + (EndBlue - StartBlue) / 255 '
    Next X
End Sub

