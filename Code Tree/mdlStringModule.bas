Attribute VB_Name = "mdlStringModule"
Option Explicit 'Declare all variables
'Enumerations for case changing
Public Enum CaseConsts
    LowerCase = 0
    UpperCase = 1
    SentenceCase = 2
    ToggleCase = 3
    TitleCase = 4
    VaryCase = 5
End Enum
'Enumerations for case changing only lower & uppercase
Public Enum UpperLowerCaseConsts
    vbLowerCase = 0
    vbUpperCase = 1
End Enum

Public Function ChangeCase(ByVal TheText As String, ByVal ChangeTo As CaseConsts, _
    Optional ByVal CaseOfFirstLetter As UpperLowerCaseConsts = UpperCase) As String

        Dim SelectedText As String
        Dim Found As Long
        Dim LastUpper As Boolean
        
        Select Case ChangeTo
            
            'lower case
            Case LowerCase
                'Return lower case thetext
                ChangeCase = LCase(TheText)
            
            'UPPER CASE
            Case UpperCase
                'Return UPPER CASE THETEXT
                ChangeCase = UCase(TheText)
            
            'Sentence case
            Case SentenceCase
                'Uppercase the first letter
                ChangeCase = UCase(Left(TheText, 1))
                'Lowercase the rest
                ChangeCase = ChangeCase & LCase(Mid(TheText, 2, Len(TheText)))
                'For Found = 3 To Len(TheText)
                    'If Mid(TheText, Found - 2, 1) = ". " Then _
                        'TheText = Left(TheText, Found) & UCase(Mid(TheText, Found, 1) & LCase(Right(TheText, Len(TheText) - (Found + 1))))
                'Next Found
            
            'tOGGLE cASE
            Case ToggleCase
                'Loop to the length of TheText
                For Found = 1 To Len(TheText)
                    'Get the letter in position of Found variable _
                     (i.e. one before current letter)
                    SelectedText = Mid(TheText, Found, 1)
                    'If it is lowercase
                    If IsLowerCase(SelectedText) = True Then
                        'Change the case and add it to ChangeCase
                        ChangeCase = ChangeCase + UCase(SelectedText)
                    'If it is uppercase
                    Else
                        'Change the case and add it to ChangeCase
                        ChangeCase = ChangeCase + LCase(SelectedText)
                    End If
                Next Found
            
            'Title Case
            Case TitleCase
                'Uppercase the first letter
                ChangeCase = UCase(Mid(TheText, 1, 1))
                'Sort the rest
                For Found = 2 To Len(TheText)
                    'If the character before the one at the variable Found is a space
                    If Mid(TheText, Found - 1, 1) = " " Then
                        'Uppercase it and add it to ChangeCase
                        ChangeCase = ChangeCase & UCase(Mid(TheText, Found, 1))
                    'If it's not a space
                    Else
                        'Lowercase it and add it to ChangeCase
                        ChangeCase = ChangeCase & LCase(Mid(TheText, Found, 1))
                    End If
                Next Found
            
            'VaRy tHe cAsE Of eAcH LeTtEr
            Case VaryCase
                'Lowercase it all
                TheText = LCase(TheText)
                'If uppercase is to be the first letter
                If CaseOfFirstLetter = vbLowerCase Then
                    'Set a variable for what the last case was
                    LastUpper = False
                    'Change it
                    TheText = LCase(Mid(TheText, 1, 1)) & Mid(TheText, 2)
                'If lowercase is wanted
                ElseIf CaseOfFirstLetter = vbUpperCase Then
                    'Set a variable for what the last case was
                    LastUpper = True
                    'Change it
                    TheText = UCase(Mid(TheText, 1, 1)) & Mid(TheText, 2)
                End If
                
                'Loop from 2 to the length of the string
                For Found = 2 To Len(TheText)
                    'If last character was uppercase
                    If LastUpper = True Then
                        'Change the current character to lowercase
                        SelectedText = LCase(Mid(TheText, Found, 1))
                    'If the last character was lowercase
                    Else
                        'Change the current character to uppercase
                        SelectedText = UCase(Mid(TheText, Found, 1))
                    End If
                    'TheText = the changed bit + the newly changed bit + the rest
                    TheText = Mid(TheText, 1, Found - 1) & SelectedText & Mid(TheText, Found + 1)
                    'Invert lastupper
                    LastUpper = Not LastUpper
                Next Found
                'Return the modified text
                ChangeCase = TheText
                
        End Select
    
End Function

Public Function CountLettersInCase(ByVal TheText As String, _
    ByVal UpperOrLowerCase As UpperLowerCaseConsts) As Long

    Dim Counter As Integer
    Dim Letter As String
    
    'Loop for the length of TheText
    For Counter = 1 To Len(TheText)
        Select Case UpperOrLowerCase
            'If the number of lowercase is wanted
            Case vbLowerCase
                'Get the current letter
                Letter = Mid(TheText, Counter, 1)
                'If it is lowercase add one to count
                If (Letter = "a") Or (Letter = "b") Or (Letter = "c") Or (Letter = "d") Or (Letter = "e") Or (Letter = "f") Or (Letter = "g") Or (Letter = "h") Or (Letter = "i") Or (Letter = "j") Or (Letter = "k") Or (Letter = "l") Or (Letter = "m") Or (Letter = "n") Or (Letter = "o") Or (Letter = "p") Or (Letter = "q") Or (Letter = "r") Or (Letter = "s") Or (Letter = "t") Or (Letter = "u") Or (Letter = "v") Or (Letter = "w") Or (Letter = "x") Or (Letter = "y") Or (Letter = "z") Then _
                    CountLettersInCase = CountLettersInCase + 1
            'If the number of uppercase is wanted
            Case vbUpperCase
                Letter = Mid(TheText, Counter, 1)
                'If it is uppercase add one to count
                If (Letter = "A") Or (Letter = "B") Or (Letter = "C") Or (Letter = "D") Or (Letter = "E") Or (Letter = "F") Or (Letter = "G") Or (Letter = "H") Or (Letter = "I") Or (Letter = "J") Or (Letter = "K") Or (Letter = "L") Or (Letter = "M") Or (Letter = "N") Or (Letter = "O") Or (Letter = "P") Or (Letter = "Q") Or (Letter = "R") Or (Letter = "S") Or (Letter = "T") Or (Letter = "U") Or (Letter = "V") Or (Letter = "W") Or (Letter = "X") Or (Letter = "Y") Or (Letter = "Z") Then _
                    CountLettersInCase = CountLettersInCase + 1
        End Select
    Next Counter
        
End Function

Public Function IsLowerCase(ByVal TheText As String) As Boolean
    
    On Error GoTo ErrorHandler
    Dim Counter As Integer
    Dim Letter As String
    
    For Counter = 1 To Len(TheText)
        Letter = Mid(TheText, Counter, 1)
        'Goto IsLetterLowerCase and find out if letter is lowercase
        GoSub IsLetterLowerCase
    Next Counter
    'Return True
    IsLowerCase = True
    'Exit function so as not cause an error
    Exit Function
    
IsLetterLowerCase:
    'If the passed variable = a to z (in lower case) the return true
    If (Letter = "A") Or (Letter = "B") Or (Letter = "C") Or (Letter = "D") Or (Letter = "E") Or (Letter = "F") Or (Letter = "G") Or (Letter = "H") Or (Letter = "I") Or (Letter = "J") Or (Letter = "K") Or (Letter = "L") Or (Letter = "M") Or (Letter = "N") Or (Letter = "O") Or (Letter = "P") Or (Letter = "Q") Or (Letter = "R") Or (Letter = "S") Or (Letter = "T") Or (Letter = "U") Or (Letter = "V") Or (Letter = "W") Or (Letter = "X") Or (Letter = "Y") Or (Letter = "Z") Then
        'Return False
        IsLowerCase = False
        Exit Function
    'If the text is not a to z (in lower case) it is therefore uppercase (or other character)
    Else
        'Go back to where we came form
        Return
    End If
    
    
ErrorHandler:
    'Tell user that the string could not be identified as lower or uppercase if they want it
    'If DisplayError = True Then _
MsgBox "Sorry, an error has occured. The specified text could not be identified as lower or uppercase text.", _
            vbExclamation + vbOKOnly, "Error"
        
End Function

Public Function IsUpperCase(ByVal TheText As String) As Boolean
    
    On Error GoTo ErrorHandler
    Dim Counter As Integer
    Dim Letter As String
    
    For Counter = 1 To Len(TheText)
        Letter = Mid(TheText, Counter, 1)
        'Goto IsLetterUpperCase and find out if letter is Uppercase
        GoSub IsLetterUpperCase
    Next Counter
    'Exit function so as not cause an error
    Exit Function
    
IsLetterUpperCase:
    'If the passed variable = A to Z (in upper case)
    If (Letter = "a") Or (Letter = "b") Or (Letter = "c") Or (Letter = "d") Or (Letter = "e") Or (Letter = "f") Or (Letter = "g") Or (Letter = "h") Or (Letter = "i") Or (Letter = "j") Or (Letter = "k") Or (Letter = "l") Or (Letter = "m") Or (Letter = "n") Or (Letter = "o") Or (Letter = "p") Or (Letter = "q") Or (Letter = "r") Or (Letter = "s") Or (Letter = "t") Or (Letter = "u") Or (Letter = "v") Or (Letter = "w") Or (Letter = "x") Or (Letter = "y") Or (Letter = "z") Then
        'Return False
        IsUpperCase = False
        'Return True
        IsUpperCase = True
        Exit Function
    'If the text is not A to Z (in upper case) it is therefore lowercase (or other character)
    Else
        'Go back to where we came from
        Return
    End If
    
    
ErrorHandler:
    'Tell user that the string could not be identified as lower or uppercase if they want it
    'If DisplayError = True Then _
        MsgBox "Sorry, an error has occured. The specified text could not be identified as lower or uppercase text.", _
            vbExclamation + vbOKOnly, "Error"
        
End Function


