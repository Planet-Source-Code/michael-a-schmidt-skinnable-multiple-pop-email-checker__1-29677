Attribute VB_Name = "modRegistry"
Option Explicit


'====================================
'   Word Function
'====================================
Public Function Word(ByVal sSource As String, n As Long, SP As String) As String
' This function is used to parse data. Data is send as
' multiple commands in one packet. Each command is seperated
' by a special character. We retrieve specific commands by
' calling 'word' and specifying what seperates each 'word'.
'=================================================
' Word retrieves the nth word from sSource
' Usage:
'    Word("red blue green ", 2)   "blue"
'=================================================
Dim pointer As Long   'start parameter of Instr()
Dim pos     As Long   'position of target in InStr()
Dim x       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

'sSource = CSpace(sSource)

'find the nth word
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If x = n Then                               'the target word-number
      lEnd = InStr(pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = Len(sSource) + 1 '   or if its the last word
      Word = Mid$(sSource, pointer, lEnd - pointer)
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
  
End Function
