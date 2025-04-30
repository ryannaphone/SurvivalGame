Attribute VB_Name = "Utilities"

Dim RawData(1 To 10000)
Dim FoundData(1 To 10000)
Global nData As Long
Global Words() As String
Global DataFileFolder As String
Global DataFileName As String
Global WordCount As Integer
Global OGWordCount As Integer
Global WordLocation As String
Global TempArray(1 To 10000) As String
Global OGIRCount As Integer
Global IRCount As Integer
Global InpRes() As String
Global Incorrect As Long
Global RandomPhrase As Integer
Global RanRes
Global UserName As String
Global UserWins As Integer
Global UserHS As Integer
'API: Allows beeps of different pitch and duration
Declare Function Beep Lib "kernel32" _
    (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'API: Play any *.wav sound
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
    (ByVal lpszName As String, ByVal hModule As Long, ByVal dwflags As Long) As Long

'API: Control mouse position
Declare Function SetCursorPos& Lib "user32.dll" (ByVal X As Long, ByVal Y As Long)

'API: mp3 player
Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal _
    lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'API: Obtain mouse position
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Global aPOINTAPI As POINTAPI

'API: Control mouse events
Declare Function mouse_event Lib "user32.dll" (ByVal dwflags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal dwData As Integer, ByVal dwExtraInfo As Integer) As Integer
Const MOUSE_EVENT_F_LEFT_DOWN As Integer = 2
Const MOUSE_EVENT_F_LEFT_UP As Integer = 4
Const MOUSE_EVENT_F_RIGHT_DOWN As Integer = 8
Const MOUSE_EVENT_F_RIGHT_UP As Integer = 16


Sub PlaySoundFile(sf$)
  'plays sound file asynchronously using API
  'sf$ is the path to the *.wav file
  PlaySound sf$, 0, &H1 Or SND_FILENAME
End Sub

Function RandomInteger(I1 As Integer, I2 As Integer) As Long
  'Return a random integer from I1 to I2
  Dim t1 As Long 'prevent overflows
  Dim t2 As Long
  Dim irange As Long
  
  If I1 > I2 Then 'force I1 < I2
    t1 = I2: t2 = I1
  Else
    t2 = I2: t1 = I1
  End If
  
  irange = t2 - t1 + 1
  
  Randomize 'randomize computer
  t = Int(Rnd() * irange) + t1
  RandomInteger = t
End Function

Public Sub MoveAndMouseClick(px As Integer, py As Integer)
  'gradually move mouse to position, then click
  Dim tx As Integer
  Dim ty As Integer
  
  For i% = 0 To 100 Step 2
    tx = Int(i% * px / 100)
    ty = Int(i% * py / 100)
    Call SetCursorPos&(tx, ty)
    DoEvents
    Call WaitForTime(0.01)
  Next i%
  
  'Create a mouse click at the final destination
  Call mouse_event(MOUSE_EVENT_F_LEFT_DOWN, px, py, 0, 0)
  Call mouse_event(MOUSE_EVENT_F_LEFT_UP, px, py, 0, 0)

End Sub

Sub WaitForTime(t)
  'Wait for t seconds, then return.
  t1 = Timer
  While Timer - t1 < t
    DoEvents
  Wend
End Sub

Sub GetMousePosition(X As Integer, Y As Integer)
  'obtain mouse position in units of screen pixels
  ret = GetCursorPos(aPOINTAPI)
  X = aPOINTAPI.X
  Y = aPOINTAPI.Y
End Sub

Function LoadTextFile(pn$, TextData$()) As Integer
  'Load lines of text into dynamic array TextData$()
  'pn$ is pathname
  'Function returns number of lines of text, or 0 if error
  On Error GoTo LoadTextFileUnsuccessful
  ReDim tData$(1 To 32767) 'temporarily array to load text

  Open pn$ For Input As #1 'open channel to text file
    nText% = 0             'count number of lines of text
    Do
      nText% = nText% + 1
      Line Input #1, tData$(nText%)
    Loop Until EOF(1) = True Or nText% = 32767
  Close #1

  ReDim TextData$(1 To nText%) 'free unneeded memory
  
  For i% = 1 To nText%
    TextData$(i%) = tData$(i%)
  Next i%
  
  LoadTextFile = nText%        'return number of lines of text
  Exit Function
  
LoadTextFileUnsuccessful: LoadTextFile = 0 'text not loaded
End Function

Function SaveTextFile(pn$, p1%, p2%, SaveText$()) As Boolean
  'Save text records p1% to p2% to file, pathname = pn$
  On Error GoTo SaveTextFileUnsuccessful

  'write text elements to file
  Open pn$ For Output As #1 'open new text file
    For i% = p1% To p2%
      Print #1, SaveText$(i%)
    Next i%
  Close #1
  
  SaveTextFile = True 'save successful
  Exit Function

SaveTextFileUnsuccessful: SaveTextFile = False 'text not saved
End Function

Function SpecialFolderPath(t) As String
 'Return the path to a special folder.
 't is one of these types:
 '    AllUsersDesktop       AllUsersStartMenu
 '    AllUsersPrograms      AllUsersStartup
 '    Desktop               Favorites
 '    Fonts                 MyDocuments
 '    NetHood               PrintHood
 '    Programs              Recent
 '    SendTo                StartMenu
 '    Startup               Templates
     
  Dim objWSHShell As Object
  Dim strSpecialFolderPath
     
  ' Create a shell object
  Set objWSHShell = CreateObject("WScript.Shell")
   ' To find out the path to the passed special folder,
   '  just change the "Desktop" for one of the other options
  SpecialFolderPath = objWSHShell.SpecialFolders(t)
   ' Clean up
  Set objWSHShell = Nothing
End Function

Function PathToDesktop() As String
  'returns path to desktop
  PathToDesktop = SpecialFolderPath("Desktop")
End Function

Sub FindX(X As String)
For i = 1 To nData
    XPos = InStr(1, RawData(i), X)
    If XPos <> 0 Then
        FoundData(i) = Mid$(RawData(i), XPos)
        'Form1.List4.AddItem FoundData(i)
        XAmount = XAmount + 1
    End If
Next i
End Sub
Sub CreateItem(ItemName As String, ItemCaption As String, ItemWidth As Long, ItemHeight As Long)
    'Load Form1.ItemName
    'Form1.ItemName.Caption = ItemCaption
    'Form1.ItemName.Width = ItemWidth
    'Form1.ItemName.Height = ItemHeight
End Sub

Sub SelectDataFile()
    Form1.CommonDialog1.FileName = ""
    Form1.CommonDialog1.DialogTitle = "Browse for Data File Folder"
    Form1.CommonDialog1.Filter = "Data Files (*.txt)|*.txt|All Files (*.*)|*.*"
    Form1.CommonDialog1.FilterIndex = 1
    Form1.CommonDialog1.Flags = &H4& Or &H8& Or &H2&
    Form1.CommonDialog1.Action = 1
    
    If Form1.CommonDialog1.FileName = "" Then
    
    Exit Sub ' if no file is selected then exit
    End If
    
    DataFileFolder$ = Form1.CommonDialog1.FileName
    DataFileName$ = Form1.CommonDialog1.FileTitle
    
    DataFileFolder$ = Left(DataFileFolder$, Len(DataFileFolder$) - Len(DataFileName$))
    Form1.CommonDialog1.InitDir = DataFileFolder$
    
    'MsgBox DataFileFolder$
    'MsgBox DataFileName$
    
End Sub

Sub LoadTextFileS(pn$, TextData$()) 'As Integer
  'Load lines of text into dynamic array TextData$()
  'pn$ is pathname
  'Function returns number of lines of text, or 0 if error
  'On Error GoTo LoadTextFileUnsuccessful
  ReDim tData$(1 To 32767) 'temporarily array to load text

  Open pn$ For Input As #2 'open channel to text file
    nText% = 0             'count number of lines of text
    Do
      nText% = nText% + 1
      Line Input #2, tData$(nText%)
    Loop Until EOF(2) = True Or nText% = 32767
  Close #2

  ReDim TextData$(1 To nText%) 'free unneeded memory
  
  For i% = 1 To nText%
    TextData$(i%) = tData$(i%)
  Next i%
  
'  LoadTextFile = nText%        'return number of lines of text
  Exit Sub
  
'LoadTextFileUnsuccessful: LoadTextFile = 0 'text not loaded
End Sub



