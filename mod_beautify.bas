Attribute VB_Name = "mod_beautify"
Option Compare Database
Option Explicit

Function beautifyModule(strModuleName As String, Optional bWhiteSpace As Boolean = True, Optional bLineNumbers = True)
' Purpose: ********************************************************************
' Formats all procedures in a given VBA module (strModuleName) with:
'   If bWhiteSpace is true, indents for loops.
'   If bWhiteSpace is false, removes all leading indents in non-comment code

'   If bLineNumbers is true, clears and re-adds line numbers to applicable rows
'   If bLineNumbers is false, clears line numbers from all non-comment rows.

' Version Control:
' Vers      Author          Date        Change
' 1.0.0     Sara Gleghorn   21/05/2020  Formats indentation for loops and line
'                                       continuations, and adds or clears line
'                                       numbers
' *****************************************************************************
' Expected Parameters:
'Dim bLineNumbers        As Boolean  ' Whether to add line numbers'
'Dim bWhiteSpace         As Boolean  ' Whether to format line spacing to make blocks of code easier to read
' In Scope Parameters:
Dim md                  As Module       ' Which modules are we in?
Dim iReadLine           As Integer      ' What line of the module are we on?
Dim strLine             As String       ' What does our line contain?
Dim strDecomment        As String       ' Line with comments stripped, to catch labels followed by comments
Dim iIndent             As Integer      ' How many levels of indent are we at?

Dim iLineNum            As Integer      ' For adding line numbers within procedure
Dim iMinIndent          As Integer      ' How many indents are minimum
Dim bAddNum             As Boolean      ' Is this line eligible for a line number?
Dim strNewLine          As String       ' For building our replacement line (with numbers/indents
Dim bLineContinuation   As Boolean      ' For managing line numbers and indent rules where " _" is used
Dim bNewIndent          As Boolean      ' Used to delay an indent if a line continuation (" _") is used on the same line as an indent-er (IF, FOR, etc)
Dim bIf                 As Boolean      ' Are we currently inside a multiline criteria statement of an IF?
Dim bExplicit           As Boolean      ' Does this module use Option Explicit?
Dim iChar               As Integer      ' For looping through characters on a line
Dim bLiteralString      As Boolean      ' For catching apostrophes and whether they are comments or in quotes
Dim bCase               As Boolean      ' Are we currently in a CASE? (Helps with aligning other CASEs)

1   On Error GoTo ErrorHandler

2   Set md = Modules(strModuleName)

UpdateModules: ' --------------------------------------------------------------
' Check for Option Explicit
3   For iReadLine = 1 To md.CountOfDeclarationLines
4       strLine = Trim(md.Lines(iReadLine, 1))
5       If strLine = "Option Explicit" Then
6           bExplicit = True
7           Exit For
8       End If
9   Next

10  If bExplicit = False Then
11      Debug.Print "Feedback: You are not using Option Explicit in module: " & md.Name _
            & vbNewLine & "    It is strongly recommended you do to protect against typos."
12  End If

13  For iReadLine = md.CountOfDeclarationLines + 1 To md.CountOfLines
       
14      strNewLine = vbNullString
15      strDecomment = vbNullString
    
16      strLine = Trim(md.Lines(iReadLine, 1))
17      If strLine = vbNullString Then GoTo ContinueFor
    
CurrentLineEffects: ' ---------------------------------------------------------
    ' Check if we're in a new Procedure
18      If md.ProcOfLine(iReadLine - 1, 0) <> md.ProcOfLine(iReadLine - 2, 0) Then
19          iLineNum = 0
20          iMinIndent = 0
21          If bLineNumbers = True Then
22              Select Case md.ProcCountLines(md.ProcOfLine(iReadLine - 1, 0), 0)
                        Case Is < 1000
24                      iMinIndent = 1
                        Case Is > 32767
26                      Err.Raise vbObjectError + 513, "beautifyModule", _
                            "One or more of the modules in " & strModuleName _
                            & "Is too long, and cannot be processed with line " _
                            & "numbers (erl is an INT and cannot exceed 32767)"
                        Case Else
28                      iMinIndent = 2
29              End Select
30          End If
        
        ' Reset current indent for this procedure
31          iIndent = iMinIndent
32      End If
        
    ' Remove the existing line number
33      If IsNumeric(Left(strLine, 1)) Then
34          strLine = Trim(Right(strLine, Len(strLine) - InStr(strLine, " ")))
35      End If

    ' Ignore firstline, blanks, comments
36      If md.ProcStartLine(md.ProcOfLine(iReadLine, 0), 0) = iReadLine - 1 Then GoTo LineContinuation
37      If strLine = vbNullString Then GoTo ContinueFor
38      If Left(strLine, 1) = "'" Then GoTo ContinueFor
39      If Left(strLine, 4) = "Dim " Then
40          strNewLine = strLine ' Trimmed
41          GoTo LineContinuation
42      End If
    
    ' Strip Comments by detecting apostophes (but not if they're within double quotes)
43      bLiteralString = False
44      For iChar = 1 To Len(strLine)
45          If Mid(strLine, iChar, 1) = """" Then
46              bLiteralString = Not bLiteralString
47          ElseIf Mid(strLine, iChar, 1) = "'" Then
48              If bLiteralString = False Then
49                  strDecomment = Trim(Left(strLine, iChar - 1))
50                  Exit For
51              End If
52          End If
53      Next
    
    ' For simplicity of logic, if we had no comments put the line in strdecomment anyway
54      If strDecomment = vbNullString Then
55          strDecomment = strLine
56      End If
    
    ' Ignore labels
57      If Right(strDecomment, 1) = ":" Then GoTo ContinueFor
        
    ' Manage Indents
58      If bLineContinuation = False Then
59          If Left(strLine, 5) = "Next " _
                Or strLine = "Next" _
                Or Left(strLine, 5) = "Loop " _
                Or strLine = "Loop" _
                Or strLine = "Else" _
                Or strLine = "End If" _
                Or Left(strLine, 7) = "ElseIf " _
                Then
60              iIndent = iIndent - 1
61          ElseIf Left(strLine, 5) = "Case " Then
62              If bCase = True Then
63                  iIndent = iIndent - 1
64              End If
65          ElseIf strLine = "End Select" Then
66              iIndent = iIndent - 2
67              bCase = False
68          End If
69      End If
    
    ' Handle if the indent goes below the minimum, which may occur if
70      If iIndent < iMinIndent Then
71          Debug.Print Now() & " beautifyModule: Attempted to reduce indent beyond the minimum. Is your syntax ok?"
72          iIndent = iMinIndent
73      End If
    
GenerateNewLine: ' ------------------------------------------------------------

74      If bWhiteSpace = False Then
75          iIndent = iMinIndent
76      End If

77      If bLineContinuation = True Then
78          If bNewIndent = True Then
79              strNewLine = Replace(Space(iIndent), " ", vbTab) & strLine
80          Else
81              strNewLine = Replace(Space(iIndent + 1), " ", vbTab) & strLine
82          End If
83      Else
84          iLineNum = iLineNum + 1
85          If bLineNumbers = True And Left(strLine, 5) <> "Case " Then
86              strNewLine = iLineNum
87              If Len(strNewLine) < 4 Then
88                  strNewLine = strNewLine & Replace(Space(iIndent), " ", vbTab)
89              Else
90                  strNewLine = strNewLine & Replace(Space(iIndent - 1), " ", vbTab)
91              End If
92          ElseIf bLineNumbers = True And Left(strLine, 5) = "Case " Then
93              strNewLine = Replace(Space(iIndent + 1), " ", vbTab)
94          Else
95              strNewLine = Replace(Space(iIndent), " ", vbTab)
96          End If
        
97          strNewLine = strNewLine & strLine
98      End If
    
FollowingLineEffects: ' -------------------------------------------------------
    ' Discover whether the next line should be indented
99      If bLineContinuation = False Then
100         If Left(strLine, 4) = "For " _
                Or Left(strLine, 3) = "Do " _
                Or Left(strLine, 7) = "ElseIf " _
                Or strLine = "Else" _
                Or Left(strLine, 12) = "Select Case " _
                Then
101             iIndent = iIndent + 1
102             bNewIndent = True
103         End If
        
104         If Left(strLine, 5) = "Case " Then
105             If bCase = False Then
106                 bCase = True
107             End If
108             iIndent = iIndent + 1
109         End If
        
110         If Left(strLine, 3) = "If " Then
111             If Right(strLine, 1) = "_" Then
112                 bIf = True
113             ElseIf Right(strDecomment, 4) = "Then" Then
114                 iIndent = iIndent + 1
115                 bIf = False
116             End If
117         End If
118     End If
    
119     If bIf = True Then
120         If Right(strDecomment, 4) = "Then" Then
121             iIndent = iIndent + 1
122             bIf = False
123         End If
124     End If
        
LineContinuation: ' -----------------------------------------------------------
125     bLineContinuation = False
126     If Right(strLine, 2) = " _" Then
127         bLineContinuation = True
128     Else
129         bNewIndent = False
130     End If
    
ReplaceLine: ' ----------------------------------------------------------------
131     If strNewLine <> vbNullString Then
132         md.ReplaceLine iReadLine, strNewLine
133     End If

ContinueFor: ' ----------------------------------------------------------------
134 Next

135 Debug.Print "Beautify Completed. Please run Debug > Compile Modules"
136 Exit Function

ErrorHandler: ' ----------------------------------------------------------------
137 Debug.Print Now() & " beautifyModule ERROR - " & Err.Number & ": " _
        & Err.Description

138 End Function
