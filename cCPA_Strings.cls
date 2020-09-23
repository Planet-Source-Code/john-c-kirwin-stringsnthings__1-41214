VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_Strings"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
  '//****************************************************************************
  '// cCPA_Strings - object to use for working with strings
  '//
  '// Version:       6.30.0
  '// Created:       12/26/2001 John C. Kirwin (JCK)
  '// Modified:      03/03/2002 JCK - formatting
  '//
  '// Dependencies:
  '//     Reference to cCPA_Tracker Object
  '//
  '// Implements:
  '//     None
  '//
  '// Example:   Dim CPAStrings As cCPA_Strings
  '//            Set CPAStrings = New cCPA_Strings
  '//
  '//**************************************************************************

Dim sMessage As String
Dim CPATracker As cCPA_Tracker                                                        ' CPATracker. Object

Public Function Alpha(sInput As String) As Boolean

  '//**************************************************************************
  '// Alpha - Verifies sInput is a letter between A - Z or a - z
  '//
  '// Created:       12/07/2002 John C. Kirwin (JCK)
  '// Modified:      12/07/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     Alpha As Boolean
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Declare local variables
    Dim lPos As Long
    Dim lLen As Long
    Dim iAscChar As Integer
    
    '//**** Get the length of the input string as lLen
    lLen = Len(sInput)
    
    '//****
    For lPos = 1 To lLen
        
        '//****
        iAscChar = Asc(UCase(Mid(sInput, lPos, 1)))
        
        '//**** Verify sInput is a number between A - Z or a - z
        If iAscChar <> 32 And iAscChar < 65 Or iAscChar > 90 Then
        
            '//****
            Exit Function
        
        '//****
        End If
        
    
    '//**** Next For i = 1 To lLen
    Next
     
    '//****
    Alpha = True

'//**** Exit Sub before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Alpha Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function

Public Function Numeric(sInput As String) As Boolean
    
  '//**************************************************************************
  '// Numeric - Verifies sInput is a number between 0 - 9
  '//
  '// Created:       12/07/2002 John C. Kirwin (JCK)
  '// Modified:      12/07/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     Numeric As Boolean
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Declare local variables
    Dim lPos As Long
    Dim lLen As Long
    Dim iAscChar As Integer
    
    '//**** Get the length of the input string as lLen
    lLen = Len(sInput)
    '//****
    For lPos = 1 To lLen
        
        '//****
        iAscChar = Asc(UCase(Mid(sInput, lPos, 1)))
        
        '//**** Verify sInput is a number between 0 - 9
        If iAscChar <> 32 And iAscChar < 48 Or iAscChar > 57 Then
        
            '//****
            Exit Function
        
        '//****
        End If
        
    
    '//**** Next For i = 1 To lLen
    Next
    
    
    '//****
    Numeric = True

'//**** Exit Sub before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Numeric Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function
Public Function AlphaNumeric(sInput As String) As Boolean

  '//**************************************************************************
  '// AlphaNumeric - Verifies sInput contains only alpha or number characters
  '//                between A - Z, a - z, or 0 - 9
  '//
  '// Created:       12/07/2002 John C. Kirwin (JCK)
  '// Modified:      12/07/2002 JCK - formatting
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     AlphaNumeric As Boolean
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Declare local variables
    Dim lPos As Long
    Dim lLen As Long
    Dim iAscChar As Integer
    
    '//**** Get the length of the input string as lLen
    lLen = Len(sInput)
    
    '//****
    For lPos = 1 To lLen
        
        '//****
        iAscChar = Asc(UCase(Mid(sInput, lPos, 1)))
        
        '//**** Verifies sInput contains only alpha or number characters
        '//     between A - Z, a - z, or 0 - 9
        If iAscChar <> 32 And (iAscChar < 65 Or iAscChar > 90) And (iAscChar < 48 Or iAscChar > 57) Then
        
            '//****
            Exit Function
        
        '//****
        End If
        
    
    '//**** Next For i = 1 To lLen
    Next
    
    
    '//****
    AlphaNumeric = True

'//**** Exit Sub before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the AlphaNumeric Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function
Function WordCount(sInput As String) As String

  '//**************************************************************************
  '// WordCnt - total words of a string assuming total equals total spaces + 1
  '//
  '// Created:       12/01/2002 John C. Kirwin (JCK)
  '// Modified:      12/01/2002 JCK - formatting
  '//
  '// Parameters:
  '//     sInput As String
  '//
  '// Returns:
  '//     WordCount as string
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Note:  UBound returns a Long containing the largest available
    '//            subscript for the indicated dimension of an array.
        WordCount = UBound(Split(sInput, " ")) + 1


'//**** Exit Sub before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the WordCount Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function

Function ReplaceString(sInput As String, sOriginal As String, sReplace As String) As String

  '//**************************************************************************
  '// ReplaceString - Replace a specified string in an input string with another string
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sInput As String - Inputstring
  '//     sOriginal As String - String to locate and replace
  '//     sReplace As String - String to replace located string with
  '//
  '// Returns:
  '//     ReplaceString as string
  '//
  '// Example:
  '//     Text1.Text = CPAStrings.ReplaceText("The input string", "string", "Text")
  '//
  '//**************************************************************************

    '//**** Error Handler
    On Error GoTo EH
  
    '//**** Original text returned if no replacement
    ReplaceString = sInput

    '//**** Declare local variables
    Dim lPos As Long
    Dim lLoc As Long

    '//**** Set start position lPos equal to 1
    lPos = 1

    '//**** Locate the first occurance of the original string
    '//     to be replaced in sInput
    lLoc = InStr(lPos, sInput, sOriginal, vbTextCompare)

    '//****
    While lLoc <> 0

        '//**** Put to gether left and right side of sInput
        '//     together with sReplace instead of sOriginal
        sInput = Left$(sInput, lLoc - 1) & sReplace & _
                 Right$(sInput, Len(sInput) - (lLoc - 1) - Len(sOriginal))

        '//**** Advance position by one character
        lPos = lLoc + 1

        '//**** Locate the next occurance of the original string
        '//     to be replaced in sInput
        lLoc = InStr(lPos, sInput, sOriginal, vbTextCompare)

    '//**** While end loop when lLoc not equal to zero
    Wend

    '//**** Finished replacing string return results
    ReplaceString = sInput

'//**** Exit Function before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ReplaceString Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Function
Public Function FillAlign(sInput As String, iLength As Integer, sString As String, _
                          Optional bAlignRight As Boolean) As String

  '//**************************************************************************
  '// FillAlign - Fill a string with a specified sting and
  '//             align left or right to length indicated
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sInput As String - input string
  '//     iLength As Integer - length
  '//     sString As String - specified character
  '//     Optional bAlignRight As Boolean - align the string right if true
  '//
  '// Returns:
  '//     FillAlign As string
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************

    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Variable Declarations
    Dim iDiff As Integer

    '//**** Evaluate if sInput is shorter than iLength
    If Len(sInput) < iLength Then

        '//**** For Next Loop through to fill sInput to acquire specified legth
        For iDiff = 1 To (iLength - Len(sInput))

            '//**** Check if left alignment indicated
            If bAlignRight Then

                '//**** Add another sString to the left of sInput
                sInput = sString & sInput

              '//**** Else left alignment indicated
              Else

                '//**** Add another sString to the right of sInput
                sInput = sInput & sString

            '//**** End if alignment check
            End If

        '//**** Continue For Next Loop to fill sInput to acquire specified legth
        Next iDiff

    '//**** End if length of sInput is shorter than iLength check
    End If

    '//**** Finished filling and aligning return results and
    '//     truncate characters if sInput is shorter than iLength
    FillAlign = Left$(sInput, iLength)

    
'//**** Exit Function before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the FillAlign Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function
Function SeparateString(sInput As String, sString As String) As String

  '//**************************************************************************
  '// SeparateString - adds a specified string between every character
  '//                  of another string
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sInput As String - input string
  '//     sString As String - string to separate with
  '//
  '// Returns:
  '//     SeparateString as string
  '//
  '// Example:
  '//     Text1.Text = CPAStrings.SeparateString(txtStartString.Text, "-")
  '//
  '//**************************************************************************
    
    '//**** Error handler
    On Error GoTo EH
    
    '//**** Local variables
    Dim iPos As Integer
    Dim sChar As String
    Dim sTemp As String
        
        '//**** Check if sInput sting is empty
        If Len(sInput) <= 0 Then
            
            '//**** No need to continue if sInput is empty
            Exit Function
        
        '//**** End If empty sInput check
        End If
        
        '//**** Check if no separator sting indicated
        If Len(sString) <= 0 Then
        
            '//**** No need to continue if no separator
            Exit Function
                
        '//**** End If no separator check
        End If
        
        '//**** For Next loop through sInput sting
        For iPos = 1 To Len(sInput)
            
            '//**** Grab next character
            sChar = Mid(sInput, iPos, 1)
            
            '//**** Build temporary string
            sTemp = sTemp & sString & sChar
        
        '//**** Continue For Next Loop
        Next
    
        '//**** Check if the first character is the separator
        If Left(sTemp, 1) = sString Then
        
            '//**** Remove the first character because its a separator
            sTemp = Right(sTemp, Len(sTemp) - 1)
        
        '//**** End if first character is separator check
        End If
        
        '//**** Finished separating temporary string
        SeparateString = sTemp
    
'//**** Exit Function before error handler
Exit Function

'//**** Error handler
EH:
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the SeparateString Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Function

Public Function Truncate(sInput As String, iLength As Integer) As String

  '//**************************************************************************
  '// Truncate - Truncates a string by a specified number of characters.
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sInput As String - The string to truncate
  '//     iLength As Integer - The number of characters to truncate
  '//
  '// Returns:
  '//     Truncate as string truncated by specified number of characters
  '//
  '// Example:
  '//     Text1.Text = CPAStrings.Truncate("The input string", 6)
  '//
  '//**************************************************************************
      
    '//**** Error Handler
    On Error GoTo EH
    
    '//**** Check length of sInput string
    If Len(sInput) <= iLength Then

        '//**** Truncate all if specified
        Truncate = ""

      '//**** Else needs to be truncated
      Else

        '//**** Grab the left of the sInput removing specified
        '//     length of characters
        Truncate = Left$(sInput, Len(sInput) - iLength)

    '//**** End If length of sInput check
    End If
   
'//**** Exit Function before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Truncate Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next

    
'//****
End Function

Function TrimString(sInput As String, sString As String, _
         bFront As Boolean, bForward As Boolean) As String

  '//**************************************************************************
  '// TrimString - trims a string before or after a certain character
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sInput As String - input string
  '//     sString As String - string to locate and replace
  '//     bFront As Boolean - Trim front true or rear if false
  '//     bForward As Boolean - Look forward true or backward false
  '//
  '// Returns:
  '//     TrimString as string
  '//
  '// Example:
  '//     Text1.Text = CPAStrings.TrimString("The input string", "string", True, True)
  '//
  '//**************************************************************************

    '//**** Error Handler
    On Error GoTo EH
  
    '//**** Original text returned if no trim
    TrimString = sInput

    '//****
    If InStr(1, sInput, sString) = False Then Exit Function

    '//**** Check search forward using InStr
    If bForward = True Then

        '//**** Check if from the front
        If bFront = True Then

            '//**** Return a specified number of characters from the left
            '//     side of a string searching from start of the string to
            '//     the end of the string for a specified string
            TrimString = Left$(sInput, InStr(1, sInput, sString, vbTextCompare) - 1)

          '//****
          Else

            '//**** Return a specified number of characters from the right
            '//     side of a string searching from end of the string to
            '//     the start of the string for a specified string
            TrimString = Right$(sInput, Len(sInput) - (InStr(1, sInput, sString)))

            '//**** End Front boolean check
        End If

      '//**** If not forward than search backward using InStrRev
      Else

        '//**** Check if from the front
        If bFront = True Then

            '//**** Return a specified number of characters from the left
            '//     side of a string searching from start of the string to
            '//     the end of the string for a specified string
            TrimString = Left$(sInput, InStrRev(sInput, sString) - 1)

          '//****
          Else

            '//**** Return a specified number of characters from the right
            '//     side of a string searching from end of the string to
            '//     the start of the string for a specified string
            TrimString = Right$(sInput, Len(sInput) - InStrRev(sInput, sString))

        '//**** End Front boolean check
        End If

    '//**** End forward boolean check
    End If

'//**** Exit Function before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the TrimString Function"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function



Private Sub Class_Initialize()

  '//**************************************************************************
  '// Class_Initialize -
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************

  
    '//**** Instantiate CPATracker object
    Set CPATracker = New cCPA_Tracker
 
'//****
End Sub


