<div align="center">

## API Console


</div>

### Description

This creates a console window through API. This is very basic, but it teaches you the basics about calling/creating a console through API. Well first:

1.) Create an exe

2.) Delete Form1

3.) Add a Module (NOT a Class Module)

4.) Insert the code below into the module

5.) Run the program and a console window should come up

I didn't document this very well because it is very basic concerning API. Feel free to use this code in your programs and also feel free to add on to this.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ravage](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ravage.md)
**Level**          |Intermediate
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ravage-api-console__1-30204/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>
<body>
<p>Option Explicit<br>
<br>
Public Declare Function AllocConsole Lib "kernel32" () As Long<br>
Public Declare Function FreeConsole Lib "kernel32" () As Long<br>
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long<br>
Public Const STD_INPUT_HANDLE = -10&amp;<br>
Public Const STD_OUTPUT_HANDLE = -11&amp;<br>
Public Const STD_ERROR_HANDLE = -12&amp;<br>
Public Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long<br>
Public Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long<br>
Public Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long<br>
Public Const FOREGROUND_BLUE = &amp;H1   ' text color contains blue.<br>
Public Const FOREGROUND_GREEN = &amp;H2   ' text color contains green.<br>
Public Const FOREGROUND_INTENSITY = &amp;H8   ' text color is intensified.<br>
Public Const FOREGROUND_RED = &amp;H4   ' text color contains red.<br>
Public Const BACKGROUND_BLUE = &amp;H10   ' background color contains blue.<br>
Public Const BACKGROUND_GREEN = &amp;H20  ' background color contains green.<br>
Public Const BACKGROUND_RED = &amp;H40   ' background color contains red.<br>
Public Const BACKGROUND_INTENSITY = &amp;H80   ' background color is intensified<br>
Public Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long<br>
Public Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long<br>
Public Const ENABLE_LINE_INPUT = &amp;H2<br>
Public Const ENABLE_ECHO_INPUT = &amp;H4<br>
Public Const ENABLE_MOUSE_INPUT = &amp;H10<br>
Public Const ENABLE_PROCESSED_INPUT = &amp;H1<br>
Public Const ENABLE_WINDOW_INPUT = &amp;H8<br>
Public Const ENABLE_PROCESSED_OUTPUT = &amp;H1<br>
Public Const ENABLE_WRAP_AT_EOL_OUTPUT = &amp;H2<br>
Private hConsoleIn As Long 'The console's input handle<br>
Private hConsoleOut As Long 'The console's output handle<br>
Private hConsoleErr As Long 'The console's error handle<br>
<br>
<br>
Private Sub Main()<br>
Dim szUserInput As String<br>
<br>
AllocConsole<br>
SetConsoleTitle "VB Console Example" 'Set the title on the console window<br>
<br>
'Get the console's handle<br>
hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)<br>
hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)<br>
hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)<br>
<br>
'Print the prompt to the user. Use the vbCrLf to get to a new line.<br>
SetConsoleTextAttribute hConsoleOut, _<br>
FOREGROUND_RED Or FOREGROUND_GREEN _<br>
Or FOREGROUND_BLUE Or FOREGROUND_INTENSITY _<br>
Or BACKGROUND_BLUE<br>
<br>
ConsolePrint "VB Console Example" &amp; vbCrLf<br>
SetConsoleTextAttribute hConsoleOut, _<br>
FOREGROUND_RED Or FOREGROUND_GREEN _<br>
Or FOREGROUND_BLUE<br>
ConsolePrint "Enter your name--> "<br>
<br>
'Get the user's name<br>
szUserInput = ConsoleRead()<br>
If Not szUserInput = vbNullString Then<br>
ConsolePrint "Hello, " &amp; szUserInput &amp; "!" &amp; vbCrLf<br>
Else<br>
ConsolePrint "Hello, whoever you are!" &amp; vbCrLf<br>
End If<br>
<br>
'End the program<br>
ConsolePrint "Press enter to exit"<br>
Call ConsoleRead<br>
<br>
FreeConsole<br>
End Sub<br>
Private Sub ConsolePrint(szOut As String)<br>
WriteConsole hConsoleOut, szOut, Len(szOut), vbNull, vbNull<br>
End Sub<br>
Private Function ConsoleRead() As String<br>
Dim sUserInput As String * 256<br>
Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)<br>
'Trim off the NULL charactors and the CRLF.<br>
ConsoleRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)<br>
End Function<br>
</p>
</body>
</html>

