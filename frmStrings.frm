VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStrings 
   Caption         =   "StringsNThings"
   ClientHeight    =   7530
   ClientLeft      =   3885
   ClientTop       =   1410
   ClientWidth     =   10815
   Icon            =   "frmStrings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10815
   Begin VB.TextBox txtStartString 
      Height          =   285
      Left            =   1560
      TabIndex        =   41
      Text            =   "The input string"
      Top             =   240
      Width           =   9015
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   40
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtResultingString 
      Height          =   285
      Left            =   1440
      TabIndex        =   32
      Top             =   7080
      Width           =   9015
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   7080
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10821
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Standard String Funtions"
      TabPicture(0)   =   "frmStrings.frx":0CCE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdProperCase"
      Tab(0).Control(1)=   "cmdUCase"
      Tab(0).Control(2)=   "cmdLCase"
      Tab(0).Control(3)=   "cmdMid"
      Tab(0).Control(4)=   "cmdLeft"
      Tab(0).Control(5)=   "cmdRight"
      Tab(0).Control(6)=   "cmdInStr"
      Tab(0).Control(7)=   "cmdStrReverse"
      Tab(0).Control(8)=   "cmdInStrRev"
      Tab(0).Control(9)=   "cmdTrim"
      Tab(0).Control(10)=   "cmdSpace"
      Tab(0).Control(11)=   "cmdLen"
      Tab(0).Control(12)=   "cmdLTrim"
      Tab(0).Control(13)=   "cmdRTrim"
      Tab(0).Control(14)=   "cmdStrComp"
      Tab(0).Control(15)=   "cmdString"
      Tab(0).Control(16)=   "Label9"
      Tab(0).Control(17)=   "lblUCase"
      Tab(0).Control(18)=   "lblLCase"
      Tab(0).Control(19)=   "lblMid"
      Tab(0).Control(20)=   "Label3"
      Tab(0).Control(21)=   "lblRight"
      Tab(0).Control(22)=   "lblInstr"
      Tab(0).Control(23)=   "lblStrReverse"
      Tab(0).Control(24)=   "lblInstrRev"
      Tab(0).Control(25)=   "lblTrim"
      Tab(0).Control(26)=   "lblSpace"
      Tab(0).Control(27)=   "lblLen"
      Tab(0).Control(28)=   "lblLTrim"
      Tab(0).Control(29)=   "lblRTrim"
      Tab(0).Control(30)=   "lblStrComp"
      Tab(0).Control(31)=   "lblString"
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Custom String Functions"
      TabPicture(1)   =   "frmStrings.frx":0CEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAlphaNumeric"
      Tab(1).Control(1)=   "cmdNumeric"
      Tab(1).Control(2)=   "cmdAlpha"
      Tab(1).Control(3)=   "cmdWordCnt"
      Tab(1).Control(4)=   "cmdFillAlign"
      Tab(1).Control(5)=   "cmdSeparateString"
      Tab(1).Control(6)=   "cmdTrimRString"
      Tab(1).Control(7)=   "cmdTruncate"
      Tab(1).Control(8)=   "cmdReplaceString"
      Tab(1).Control(9)=   "cmdTrimLString"
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(12)=   "Label6"
      Tab(1).Control(13)=   "Label1"
      Tab(1).Control(14)=   "Label2"
      Tab(1).Control(15)=   "lblSeparateString"
      Tab(1).Control(16)=   "lblTrimRString"
      Tab(1).Control(17)=   "lblTruncate"
      Tab(1).Control(18)=   "lblReplaceString"
      Tab(1).Control(19)=   "lblTrimString"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Custom Parsing Functions"
      TabPicture(2)   =   "frmStrings.frx":0D06
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdParseString"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdLoadArray"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdJoin"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdJoin 
         Caption         =   "Join"
         Height          =   300
         Left            =   120
         TabIndex        =   63
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton cmdProperCase 
         Caption         =   "ProperCase"
         Height          =   300
         Left            =   -74880
         TabIndex        =   61
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlphaNumeric 
         Caption         =   "AlphaNumeric"
         Height          =   300
         Left            =   -74880
         TabIndex        =   59
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdNumeric 
         Caption         =   "Numeric"
         Height          =   300
         Left            =   -74880
         TabIndex        =   57
         Top             =   3300
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlpha 
         Caption         =   "Alpha"
         Height          =   300
         Left            =   -74880
         TabIndex        =   55
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoadArray 
         Caption         =   "TotalElements"
         Height          =   300
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdParseString 
         Caption         =   "ParseString"
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton cmdWordCnt 
         Caption         =   "WordCount"
         Height          =   300
         Left            =   -74880
         TabIndex        =   49
         Top             =   2700
         Width           =   1215
      End
      Begin VB.CommandButton cmdFillAlign 
         Caption         =   "FillAlign"
         Height          =   300
         Left            =   -74880
         TabIndex        =   48
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdSeparateString 
         Caption         =   "SeparateString"
         Height          =   300
         Left            =   -74880
         TabIndex        =   45
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdTrimRString 
         Caption         =   "TrimRString"
         Height          =   300
         Left            =   -74880
         TabIndex        =   43
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton cmdTruncate 
         Caption         =   "Truncate"
         Height          =   300
         Left            =   -74880
         TabIndex        =   36
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton cmdReplaceString 
         Caption         =   "ReplaceString"
         Height          =   300
         Left            =   -74880
         TabIndex        =   35
         Top             =   2100
         Width           =   1215
      End
      Begin VB.CommandButton cmdTrimLString 
         Caption         =   "TrimLString"
         Height          =   300
         Left            =   -74880
         TabIndex        =   34
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdUCase 
         Caption         =   "UCase"
         Height          =   300
         Left            =   -74880
         TabIndex        =   15
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton cmdLCase 
         Caption         =   "LCase"
         Height          =   300
         Left            =   -74880
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdMid 
         Caption         =   "Mid"
         Height          =   300
         Left            =   -74880
         TabIndex        =   13
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "Left"
         Height          =   300
         Left            =   -74880
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   "Right"
         Height          =   300
         Left            =   -74880
         TabIndex        =   11
         Top             =   2100
         Width           =   1215
      End
      Begin VB.CommandButton cmdInStr 
         Caption         =   "Instr"
         Height          =   300
         Left            =   -74880
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdStrReverse 
         Caption         =   "StrReverse"
         Height          =   300
         Left            =   -74880
         TabIndex        =   9
         Top             =   2700
         Width           =   1215
      End
      Begin VB.CommandButton cmdInStrRev 
         Caption         =   "InstrRev"
         Height          =   300
         Left            =   -74880
         TabIndex        =   8
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdTrim 
         Caption         =   "Trim"
         Height          =   300
         Left            =   -74880
         TabIndex        =   7
         Top             =   4500
         Width           =   1215
      End
      Begin VB.CommandButton cmdSpace 
         Caption         =   "Space"
         Height          =   300
         Left            =   -74880
         TabIndex        =   6
         Top             =   3300
         Width           =   1215
      End
      Begin VB.CommandButton cmdLen 
         Caption         =   "Len"
         Height          =   300
         Left            =   -74880
         TabIndex        =   5
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdLTrim 
         Caption         =   "LTrim"
         Height          =   300
         Left            =   -74880
         TabIndex        =   4
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdRTrim 
         Caption         =   "RTrim"
         Height          =   300
         Left            =   -74880
         TabIndex        =   3
         Top             =   5100
         Width           =   1215
      End
      Begin VB.CommandButton cmdStrComp 
         Caption         =   "StrComp"
         Height          =   300
         Left            =   -74880
         TabIndex        =   2
         Top             =   3900
         Width           =   1215
      End
      Begin VB.CommandButton cmdString 
         Caption         =   "String"
         Height          =   300
         Left            =   -74880
         TabIndex        =   1
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Join array elements into a string - CPAParse.JoinElements"
         Height          =   300
         Left            =   1440
         TabIndex        =   64
         Top             =   1550
         Width           =   9000
      End
      Begin VB.Label Label9 
         Caption         =   "Converts a string to proper case - StrConv(""The input string"", vbProperCase)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   62
         Top             =   5450
         Width           =   9000
      End
      Begin VB.Label Label8 
         Caption         =   "Verifies a string contains only alpha or number characters - CPAStrings.AlphaNumeric(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   60
         Top             =   3645
         Width           =   9000
      End
      Begin VB.Label Label7 
         Caption         =   "Verifies a string is a number between 0 - 9 - CPAStrings.Numeric(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   58
         Top             =   3345
         Width           =   9000
      End
      Begin VB.Label Label6 
         Caption         =   "Verifies a string is a letter between A - Z or a - z - CPAStrings.Alpha(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   56
         Top             =   3045
         Width           =   9000
      End
      Begin VB.Label Label5 
         Caption         =   "Return the total number of members of the array as Long - CPAParseDelimited.LoadArray(""The input string"", "" "")"
         Height          =   300
         Left            =   1440
         TabIndex        =   54
         Top             =   1245
         Width           =   8640
      End
      Begin VB.Label Label4 
         Caption         =   "Parses out a field as a string from a delimited string - CPAParseDelimited.ParseString(""The input string"", 2, "" "")"
         Height          =   300
         Left            =   1440
         TabIndex        =   52
         Top             =   950
         Width           =   9000
      End
      Begin VB.Label Label1 
         Caption         =   "Counts total words of a string assuming total equals total spaces + 1 - CPAStrings.WordCount(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   50
         Top             =   2745
         Width           =   9015
      End
      Begin VB.Label Label2 
         Caption         =   "Fill or truncate a string to a specified length and alignment - CPAStrings.FillAlign(txtStartString.Text, 20, Space$(1), True) "
         Height          =   300
         Left            =   -73560
         TabIndex        =   47
         Top             =   2445
         Width           =   9015
      End
      Begin VB.Label lblSeparateString 
         Caption         =   "Separates a string with a specified string - CPAStrings.SeparateString(""The input string"", ""-"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   46
         Top             =   1850
         Width           =   9000
      End
      Begin VB.Label lblTrimRString 
         Caption         =   "Trims the rear off a string after a specified character - CPAStrings.TrimString(""The input string"", Space$(1), False, False)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   44
         Top             =   1545
         Width           =   9015
      End
      Begin VB.Label lblTruncate 
         Caption         =   "Truncates a string by a specified number of characters - CPAStrings.Truncate(""The input string"", 6)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   39
         Top             =   950
         Width           =   9000
      End
      Begin VB.Label lblReplaceString 
         Caption         =   "Replace a specified string in an input string with another string - CPAStrings.ReplaceString(""The input string"", ""string"", ""Text"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   38
         Top             =   2150
         Width           =   9000
      End
      Begin VB.Label lblTrimString 
         Caption         =   "Trims the front off a string before a specified character - CPAStrings.TrimString(""The input string"", Space$(1), True, True)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   37
         Top             =   1245
         Width           =   9000
      End
      Begin VB.Label lblUCase 
         Caption         =   "Returns the specified string, converted to uppercase - UCase(""The Input String"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   30
         Top             =   950
         Width           =   9000
      End
      Begin VB.Label lblLCase 
         Caption         =   "Returns the specified string, converted to lowercase - LCase(""The Input String"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   29
         Top             =   1250
         Width           =   9000
      End
      Begin VB.Label lblMid 
         Caption         =   "Returns specified number of characters from a string - Mid(""The input string"", 2, 5)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   28
         Top             =   1550
         Width           =   9000
      End
      Begin VB.Label Label3 
         Caption         =   "Returns a specified number of characters from the left side of a string - Left(""The input string"", 5)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   27
         Top             =   1850
         Width           =   9000
      End
      Begin VB.Label lblRight 
         Caption         =   "Returns a specified number of characters from the right side of a string - Right(""The input string"", 5)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   26
         Top             =   2150
         Width           =   9000
      End
      Begin VB.Label lblInstr 
         Caption         =   $"frmStrings.frx":0D22
         Height          =   300
         Left            =   -73560
         TabIndex        =   25
         Top             =   2450
         Width           =   9000
      End
      Begin VB.Label lblStrReverse 
         Caption         =   "Reverse a string - StrReverse(""The input string"".Text)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   24
         Top             =   2750
         Width           =   9000
      End
      Begin VB.Label lblInstrRev 
         Caption         =   "Returns the position of the last occurrence of one string within another - InStrRev(""The input sting"", ""string"", , vbTextCompare)"
         Height          =   300
         Left            =   -73560
         TabIndex        =   23
         Top             =   3050
         Width           =   9000
      End
      Begin VB.Label lblTrim 
         Caption         =   "Returns a copy of a string without leading and trailing spaces - Trim(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   22
         Top             =   4550
         Width           =   9000
      End
      Begin VB.Label lblSpace 
         Caption         =   "Returns a string consisting of the specified number of spaces - Space(5) + ""The input string"""
         Height          =   300
         Left            =   -73560
         TabIndex        =   21
         Top             =   3350
         Width           =   9000
      End
      Begin VB.Label lblLen 
         Caption         =   "Returns string length or bytes required to store a variable - Len(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   20
         Top             =   3650
         Width           =   9000
      End
      Begin VB.Label lblLTrim 
         Caption         =   "Returns a copy of a string without leading spaces - LTrim(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   19
         Top             =   4850
         Width           =   9000
      End
      Begin VB.Label lblRTrim 
         Caption         =   "Returns a copy of a string without trailing spaces - RTrim(""The input string"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   18
         Top             =   5150
         Width           =   9000
      End
      Begin VB.Label lblStrComp 
         Caption         =   "Returns the result of a string comparison - StrComp(""The input string"", sText, vbTextCompare)"
         Height          =   255
         Left            =   -73560
         TabIndex        =   17
         Top             =   3950
         Width           =   9000
      End
      Begin VB.Label lblString 
         Caption         =   "Returns a repeating character string of the length specified - String(Len(""The input text""), ""*"")"
         Height          =   300
         Left            =   -73560
         TabIndex        =   16
         Top             =   4250
         Width           =   9000
      End
   End
   Begin VB.Label lblStartingString 
      Caption         =   "Starting String"
      Height          =   300
      Left            =   1560
      TabIndex        =   42
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblResultingString 
      Caption         =   "Resulting String"
      Height          =   300
      Left            =   1440
      TabIndex        =   33
      Top             =   6840
      Width           =   1455
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  '//**************************************************************************
  '// StringsNThings - Standard string functions demo with some custom string
  '//                  functions also included
  '//
  '// Version:       6.30.0
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '// Modified:      12/07/2002 JCK - added parsing by delimiter functions
  '//
  '// Standard References:
  '//     Visual Basic For Applications
  '//     Visual Basic runtime objects and procedures
  '//     Visual Basic objects and procedures
  '//     OLE Automation
  '//
  '// Custom References:
  '//     cCPA_Files - Custom file functions
  '//     cCPA_Strings - Custom string functions
  '//     cCPA_Tracker - Log and track events and errors
  '//     cCPA_Parse - Custom parsing functions
  '//
  '// Standard Components:
  '//     Microsoft Tabbed Dialog Control v6.0 (SP5)
  '//
  '// Custom Components:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '// Feedback & Enhancement Requests:
  '//     12/6/2002 John don't bother with UI, but fill that custom
  '//               string section a bit, and you 'll get your 5.
  '//     12/6/2002 Add something like change delimiter, and make
  '//               something like a wizard and you 'll deserve 5.
  '//
  '//
  '//
  '//
  '//**************************************************************************

'//**** Global Variables
Dim sMessage As String

'//**** Object Declarations
Dim CPAParse As cCPA_Parse                                 ' CPAParse Object
Dim CPAStrings As cCPA_Strings                                               ' CPAStrings Object
Dim CPATracker As cCPA_Tracker                                               ' CPATracker Object

Private Sub cmdJoin_Click()

  '//**************************************************************************
  '// Join - Join array elements into a string
  '//
  '// Created:       12/07/2002 John C. Kirwin (JCK)
  '// Modified:      12/07/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     Function Join(SourceArray, [Delimiter]) As String
  '//
  '//**************************************************************************
  
    Dim x As Long
    
    '//**** Error Handler
    On Error GoTo EH
    
    
    '//**** Use LoadArray function of CPAParse object to load the array
    '//     with the string and return the total number of members of
    '//     the array as Long and enter result in txtResultingString.Text
    x = CPAParse.LoadArray(txtStartString.Text, " ")
    
    
    '//**** Use JoinElements function of CPAParse object to join the
    '//     the members of the array as string and enter result
    '//     in txtResultingString.Text
    txtResultingString.Text = CPAParse.JoinElements
    


    
    
'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Join Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdLoadArray_Click()
  
  '//**************************************************************************
  '// LoadArray - Return the total number of members of the array as Long
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
  
  
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use LoadArray function of CPAParse object to return the
    '//     total number of members of the array as Long and enter result
    '//     in txtResultingString.Text
    txtResultingString.Text = CPAParse.LoadArray(txtStartString.Text, " ")

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ParseString Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdAlpha_Click()
  
  '//**************************************************************************
  '// Alpha - Verifies a string is a letter between A - Z or a - z
  '//
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
  
  
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use Alpha function of CPAStrings object to verify a string
    '//     is a letter between A - Z or a - z and enter result in txtResultingString.Text
    txtResultingString.Text = CPAStrings.Alpha(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Alpha Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub
Private Sub cmdNumeric_Click()
  
  '//**************************************************************************
  '// Numeric - Verifies a string is a number between 0 - 9
  '//
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
  
  
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use Numeric function of CPAStrings object to verify a string
    '//     is numeric and enter result in txtResultingString.Text
    txtResultingString.Text = CPAStrings.Numeric(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Numeric Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdAlphaNumeric_Click()
  
  '//**************************************************************************
  '// AlphaNumeric - Verifies a string contains only alpha or number characters
  '//                between A - Z, a - z, or 0 - 9
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
  
  
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use AlphaNumericfunction of CPAStrings object to verify a string
    '//     is alpha or numeric and enter result in txtResultingString.Text
    txtResultingString.Text = CPAStrings.AlphaNumeric(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the AlphaNumeric Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub












Private Sub cmdParseString_Click()
  '//**************************************************************************
  '// ParseString - Parses out a field as a string from a delimited string
  '//               sString ByVal so changes are not propagated back to the caller
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
  
  
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use ParseString function of CPAStrings object to parses out
    '//     a specified string within another delimited string and
    '//     enter result in txtResultingString.Text
    txtResultingString.Text = CPAParse.ParseString(txtStartString.Text, _
                                                        2, " ")

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ParseString Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdReplaceString_Click()

  '//**************************************************************************
  '// ReplaceString - Replace a specified string within another string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use ReplaceString function of CPAStrings object to replace
    '//     a specified string within another string and enter result
    '//     in txtResultingString.Text
    txtResultingString.Text = CPAStrings.ReplaceString(txtStartString.Text, _
                                                       "string", "Text")

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ReplaceString Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub





Private Sub cmdTruncate_Click()

  '//**************************************************************************
  '// Truncate - truncate a specified number of characters from end of string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Truncate 6 characters from end of txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = CPAStrings.Truncate(txtStartString.Text, 6)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Truncate Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub
Private Sub cmdFillAlign_Click()

  '//**************************************************************************
  '// FillAlign - fill or truncate a string to a specified length and alignment
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
    
    '//**** Error Handler
    On Error GoTo EH

    '//**** Use FillAlign function of CPAStrings object to fill
    '//     txtStartString.Text to a length of 20 characters aligned
    '//     to the right and enter result in txtResultingString.Text
    '//
    '//     Note the use of the Space$(1) function to indicate a single
    '//     space could have been any character
    txtResultingString.Text = CPAStrings.FillAlign(txtStartString.Text, 20, Space$(1), True)
    
'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the FillAlign Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

'//****
End Sub

Private Sub cmdSeparateString_Click()

  '//**************************************************************************
  '// TrimRString - separate a string with a specified string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use SeparateString function of CPAStrings object to separate the
    '//     characters of txtStartString.Text with a dash and enter
    '//     result in txtResultingString.Text
    txtResultingString.Text = CPAStrings.SeparateString(txtStartString.Text, "-")
                                                        
'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the SeparateString Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

'//****
End Sub

Private Sub cmdTrimRString_Click()

  '//**************************************************************************
  '// TrimRString - trim string from the right of the first occurance of
  '//               specified string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
    
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use TrimString function of CPAStrings object to trim the
    '//     characters to the right of the last space in txtStartString.Text
    '//     and enter result in txtResultingString.Text
    '//
    '//     Note the use of the Space$(1)function to indicate a single
    '//     space could have been any character
    txtResultingString.Text = CPAStrings.TrimString(txtStartString.Text, _
                                                    Space$(1), False, False)
'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the TrimString Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub
Private Sub cmdTrimLString_Click()

  '//**************************************************************************
  '// TrimLString - trim string from the left of the first occurance of
  '//               specified string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use TrimString function of CPAStrings object to trim the
    '//     characters to the left of the first space in txtStartString.Text
    '//     and enter result in txtResultingString.Text
    '//
    '//     Note the use of the Space$(1)function to indicate a single space
    '//     could have been any character
    txtResultingString.Text = CPAStrings.TrimString(txtStartString.Text, _
                                                    Space$(1), True, True)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the TrimLString Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub
Private Sub cmdWordCnt_Click()

  '//**************************************************************************
  '// WordCount - total words of a string assuming total equals total spaces + 1
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Use WordCountfunction of CPAStrings object to count the total
    '//     number of words in a the txtStartString.Text and enter
    '//     result in txtResultingString.Text
    txtResultingString.Text = CPAStrings.WordCount(txtStartString.Text)
                                                        
'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the WordCount Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

'//****
End Sub

Private Sub cmdUCase_Click()

  '//**************************************************************************
  '// UCase$ - upper case the whole string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Convert stringtxtStartString.Text to upper case
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = UCase$(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the UCase function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Sub

Private Sub cmdSpace_Click()

  '//**************************************************************************
  '// Space$ - add spaces to another string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Add 5 spaces to txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = Space$(5) & txtStartString.Text

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Space function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdLen_Click()

  '//**************************************************************************
  '// Len - returns the length of the string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Return the length of txtStartString.Text and enter result
    '//     in txtResultingString.Text
    txtResultingString.Text = Len(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Len function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Sub
Private Sub cmdStrComp_Click()

  '//**************************************************************************
  '// StrComp - string comparison as follows:
  '//           string1 is less than string2 then -1 is returned
  '//           string1 is equal to string2 then 0 is returned
  '//           string1 is greater than string2 then 1 is returned
  '//           string1 or string2 is Null then a Null value is returned
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Compare txtStartString.Text to "The input string"
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = StrComp(txtStartString.Text, "The input string", _
                                      vbTextCompare)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the StrComp function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdString_Click()

  '//**************************************************************************
  '// String$ - repeating character string of the length specified
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Repeat * character string of the length of txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = String$(Len(txtStartString.Text), "*")

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the String function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Sub
Private Sub cmdProperCase_Click()
  '//**************************************************************************
  '// ProperCase - Converts a string to proper case
  '//
  '// Created:       12/07/2002 John C. Kirwin (JCK)
  '// Modified:      12/07/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     StrConv("The input string", vbProperCase)
  '//
  '//**************************************************************************
  
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Converts the string txtStartString.Text to proper case
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = StrConv(txtStartString.Text, vbProperCase)
    
'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ProperCase Sub"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdLCase_Click()

  '//**************************************************************************
  '// LCase$ - lower case specified string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Lower case txtStartString.Text and enter result in
    '//     txtResultingString.Text
    txtResultingString.Text = LCase$(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the LCase function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdMid_Click()

  '//**************************************************************************
  '// Mid$ - return part of a string from any part of another string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     mid("the String", Starting position, Ending Position)
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Return 5 characters from txtStartString.Text starting at second
    '//     character and enter result in txtResultingString.Text
    txtResultingString.Text = Mid$(txtStartString.Text, 2, 5)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Mid function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdLeft_Click()

  '//**************************************************************************
  '// Left$ - returns specified length of characters from the left side
  '//         of specified sting
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Returns 5 characters from the left side of txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = Left$(txtStartString.Text, 5)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Left function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Sub

Private Sub cmdRight_Click()

  '//**************************************************************************
  '// Right$ - returns specified length of characters from the right side
  '//          of specified sting
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Returns 5 characters from the right side of txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = Right$(txtStartString.Text, 5)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Right function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub
Private Sub cmdInStr_Click()

  '//**************************************************************************
  '// InStr - returns starting position of a specified string within another string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Returns starting position of "string" within txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = InStr(1, txtStartString.Text, "string", _
                                    vbTextCompare)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the InStr function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Sub
Private Sub cmdStrReverse_Click()

  '//**************************************************************************
  '// StrReverse - reverses direction of specified string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Reverse direction of txtStartString.Text and enter result
    '//     in txtResultingString.Text
    txtResultingString.Text = StrReverse(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the StrReverse function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdInStrRev_Click()

  '//**************************************************************************
  '// InStrRev - search for specified string in another string starting
  '//            from the end of string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     InStrRev("The string to be searched", "search", _
  '//              starting position, type of search)
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Search for "string" in txtStartString.Text starting from the
    '//     end and enter result in txtResultingString.Text
    txtResultingString.Text = InStrRev(txtStartString.Text, "string", _
                                       , vbTextCompare)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the InStrRev function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdTrim_Click()

  '//**************************************************************************
  '// cmdTrim_Click - removes spaces from front and back of specified string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Remove all spaces from front and back of txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = Trim$(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Trim function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub
Private Sub cmdLTrim_Click()

  '//**************************************************************************
  '// LTrim$ - trim the spaces from the left side of the string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Trim the spaces from the left side of txtStartString.Text
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = LTrim$(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the LTrim function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

Private Sub cmdRTrim_Click()

  '//**************************************************************************
  '// RTrim$ - trim the spaces off the right side of the string
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Trim the spaces off the right side of the string
    '//     and enter result in txtResultingString.Text
    txtResultingString.Text = RTrim$(txtStartString.Text)

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the RTrim function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub






Private Sub Form_Load()

  '//**************************************************************************
  '// Form_Load - Instantiate objects and set global variable values
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH

    '//**** Enter demo string into txtStartString textbox
    txtStartString.Text = "The input string"
    
    '//**** Clear txtResultingString.Text
    txtResultingString.Text = ""
    
    '//**** Instantiate objects
    Set CPAParse = New cCPA_Parse
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker

'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during Form Load"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub
Private Sub cmdReset_Click(Index As Integer)
   
  '//**************************************************************************
  '// GetStartSting - enters demo string into txtStartString textbox
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     Index As Integer indicating which Reset command button of the
  '//     control array clicked
  '//
  '// Returns:
  '//     None
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
   
    '//**** Enter demo string into txtStartString textbox
    txtStartString.Text = "The input string"

    '//**** Clear txtResultingString.Text
    txtResultingString.Text = ""
    
'//**** Exit Sub before error handler
Exit Sub

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Reset Click Event"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Sub

