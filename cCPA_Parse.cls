VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_Parse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//****************************************************************************
'// cCPA_Parse - Class Interface to delimited string parcing utilities
'//
'// Version:       6.20.0
'// Created:       07/28/2002 John C. Kirwin (JCK)
'// Modified:      07/28/2002 JCK - Format & Comment
'//
  '// Dependencies:
  '//     Reference to cCPA_Tracker Object
  '//
  '// Implements:
  '//     None
  '//
  '// Example:   Dim CPAParse As cCPA_Parse
  '//            Set CPAParse = New cCPA_Parse
  '//
  '//**************************************************************************

'//**** Local variables to hold property values.
Private sMessage As String
Private mStrArray() As String

Private CPATracker As cCPA_Tracker                                                        ' CPATracker Object
Public Function LoadArray(ByVal sInput As String _
                        , ByVal sDelimiter As String) As Long

  '//**********************************************************************
  '// LoadArray - Loads string into an array and returns the total number
  '//             of members of the array as Long
  '//
  '// Version:  6.30.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '// ByVal sInput As String
  '// ByVal lDelimiter As Long
  '//
  '// Returns:
  '//     LoadArray As Long
  '//
  '// Example:
  '//     LoadArray("The input string", " ")
  '//       mStrArray(0)
  '//       mStrArray(1) - The
  '//       mStrArray(2) - input
  '//       mStrArray(3) - string
  '//
  '//**************************************************************************

    '//**** Error Handler
    On Error GoTo EH

    '//**** Declare variables
    Dim lPos As Long
    Dim lNextPos As Long
    Dim lArrayMemberCount As Long
    Dim varValue As Variant
  
    '//**** Wrap sInput with sDelimiter
    sInput = sDelimiter & sInput & sDelimiter

    '//**** Initialize the count and position
    lArrayMemberCount = 0
    
    '//**** Position after inserted delimeter
    lPos = Len(sDelimiter) + 1
    
    '//**** Next Position
    lNextPos = InStr(lPos, sInput, sDelimiter)

    '//**** Loop searching for delimiters
    Do Until lNextPos = 0

        '//**** Extract a new member to store in the mStrArray array
        varValue = Mid$(sInput, lPos, lNextPos - lPos)

        '//**** Add room for the new member in the array
10      ReDim Preserve mStrArray(lArrayMemberCount) As String

        '//**** Put the member data in the array
        mStrArray(lArrayMemberCount) = varValue
        
        '//**** Increment the data counter
        lArrayMemberCount = lArrayMemberCount + 1

        '//**** Check if lNextPos has reached the end of the string
        If lNextPos = 0 Then
            
            '//**** No more delimiters so exit Do Loop
            Exit Do
          
        '//**** If not keep going
        Else

            '//**** Position to the last found delimiter
            lPos = lNextPos + Len(sDelimiter)

            '//**** Find the next delimiter
            lNextPos = InStr(lPos, sInput, sDelimiter)

            '//**** Check if lPos is greater than the length of the string
            If lPos >= Len(sInput) Then
                
                '//****
                varValue = ""
                
                '//****
                lNextPos = 0
                
                '//****
                GoTo 10
            
            '//**** End if lPos check
            End If
        
        '//****
        End If

    '//****
    Loop
    
    '//**** Remove the last element of the array because its empty
    ReDim Preserve mStrArray(lArrayMemberCount - 1) As String
    
    '//**** Return the total number of members of the array
    LoadArray = UBound(mStrArray) '(lArrayMemberCount - 2)


'//**** Exit Sub/Function before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the LoadArray function"
    
    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function
Public Property Get Count() As Long

  '//**********************************************************************
  '// Count - Return largest element of the array
  '//
  '// Version:  6.30.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
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

    '//**** Return largest element of the array
    Count = UBound(mStrArray)

'//****
End Property
Public Property Get JoinElements() As String

  '//**********************************************************************
  '// JoinElements - Join all the elements of an array
  '//
  '// Version:  6.30.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
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

    '//**** Join all the elements of an array
    JoinElements = Join(mStrArray)

'//****
End Property
Public Function ParseString(ByVal sString As String _
                            , ByVal iNumber As Integer _
                            , sDelimiter As String) As String

  '//**********************************************************************
  '// ParseString - Parses out a field as a string from a delimited string
  '// sString ByVal so changes are not propagated back to the caller
  '//
  '// Version:  6.20.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '// sString As String - delimited string
  '// iNumber As Integer - number of the field to be parsed out of delimited string
  '// sDelimiter As String - Delimiter
  '//
  '// Returns:
  '//     ParseString - string parsed out of a delimited string
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************

    '//**** Error Handler
    On Error GoTo EH

    '//**** Variable Declarations
  Dim iFieldPos As Integer
  Dim iFieldPosold As Integer
  Dim iCurrentPosition As Integer
  Dim sText As String

    '//**** Wrap sString with delimeters
    sString = sDelimiter & sString & sDelimiter

    '//**** Verify there is a string to parce sString and that it conatains the
    '//     delimeter as string sDelimeter and the section to be parsed is
    '//     indicated by an integer greater than 0
    If Len(sString) > 0 And InStr(sString, sDelimiter) > 0 And iNumber > 0 Then

        '//****
        iFieldPos = 1

        '//****
        iFieldPosold = 1

        '//**** Validate the field to parse out of the string indicated
        '//     by the iNumber is not greater than the total number of
        '//     the fields separated by the sDelimiter
        Do While InStr(iFieldPosold + 1, sString, sDelimiter) > 0

            '//**** Track the old field position
            iFieldPosold = iFieldPos

            '//**** Return the position of the first occurrence of the string
            iFieldPos = InStr(iFieldPos + 1, sString, sDelimiter)

            '//**** Increment Current Position by 1
            iCurrentPosition = iCurrentPosition + 1

            '//**** Continue Do While...Loop through fields of sString
            '//     separated by the sDelimiter
        Loop

        '//**** Compare the the field to parse out is greater than
        '//     the total parsible fields
        If iNumber > iCurrentPosition Then

            '//**** Exit function if it is greater than the total
            '//     parsible fields
            Exit Function

        End If

        '//**** Reset field position integer iFieldPos to 1
        iFieldPos = 1

        '//**** Reset current position integer iCurrentPosition to 0
        iCurrentPosition = 0

        '//**** Begin the procedure to parse out the iNumber field
        '//     from sString
        Do

            '//**** Track field position
            iFieldPosold = iFieldPos

            '//****
            iFieldPos = InStr(iFieldPos + 1, sString, sDelimiter)

            '//****
            If iFieldPos > 0 Then

                '//****
                sText = Mid$(sString, iFieldPosold + 1, (iFieldPos - 1) - iFieldPosold)

              Else
                '//****
                sText = Mid$(sString, iFieldPosold + 1)

            End If

            '//**** Increment Current Position by 1
            iCurrentPosition = iCurrentPosition + 1

            '//**** Compare the current position integer iCurrentPosition
            '//     to the field integer indicated as the field to parse
            '//     out of the string sString
            If iCurrentPosition = iNumber Then

                '//**** Set ParseString string function
                ParseString = sText

                Exit Do

            End If

            '//**** Continue Do...Loop
        Loop

    End If

    '//****
    ParseString = sText

    '//**** Exit Sub/Function before error handler

Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ParseString function"
    
    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function
Public Function sParse(sBuf As String, sDelim As String) As String

  '//**************************************************************************
  '// sParse - Returns a section of a string from the beginning to the
  '//          first occurance of a delimitter.
  '// Version:  6.20.0
  '// Created:  03/03/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sBuf As String - (in/out) - The string to parse
  '//     sDelim As String - The delimitter used for the parse
  '//
  '// Returns:
  '//     sParse as string parsed from the beginning a sBuf string
  '//     to the first occurance of the sDelim delimitter indicated
  '//
  '// Example:
  '//     None
  '//
  '//**************************************************************************

    '//**** Error Handler
    On Error GoTo EH
    
  '//****
  Dim iPosition As Integer

    '//****
    iPosition = InStr(sBuf, sDelim)

    '//**** If delimitter does not exist in string, return whole string
    If iPosition = 0 Then

        '//****
        sParse = Trim$(sBuf)

        '//****
        sBuf = ""

        '//****
      Else

        '//****
        sParse = Trim$(Left$(sBuf, iPosition - 1))

        '//****
        sBuf = Mid$(sBuf, iPosition + Len(sDelim))

        '//****
    End If

'//**** Exit Sub/Function before error handler
Exit Function

'//**** Error handler
EH:
    
    '//**** Capture error information
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the sParse function"
    
    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function
    Resume Next
    
'//****
End Function
Private Sub Class_Initialize()

  '//**************************************************************************
  '// Class_Initialize -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
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

  '//****
  'Set CPATracker = New cCPA_Tracker

End Sub


