Attribute VB_Name = "mdlFileFunctions"
Option Explicit

Public Function OpenText(ByVal Filename As String) As String
    
    On Error GoTo ErrorHandler
    Dim FileNumber As Integer
    Dim TempText As String
                                        
    FileNumber = FreeFile
    'Open the file for input
    Open Filename For Input As #FileNumber
        'Return the file's contents
        OpenText = Input(LOF(FileNumber), FileNumber)
    'Close the file
    Close #FileNumber
    
    
    'Exit the function so as not casue an error
    Exit Function
    
ErrorHandler:

End Function

Public Function DoesFileExist(ByVal Filename As String) As Boolean

    On Error GoTo ErrorHandler
    Dim FileNumber As Integer
    
    FileNumber = FreeFile
    'Open the file - if it exists no error will occur and continue to next statement
    Open Filename For Input As #FileNumber
    'Close it
    Close #FileNumber
    'return true if the length of the file is > 0
    If Len(Dir$(Filename)) > 0 Then
        DoesFileExist = True
        Exit Function
    Else
        DoesFileExist = False
    End If
    
ErrorHandler:
    'Close the file
    Close #FileNumber
    'return False
    DoesFileExist = False
End Function
