VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author: Le Duc Hong         http://www.vovisoft.com
Option Explicit
Private strText As String      ' the whole string
Private strDelimiter As String ' the delimiter,  eg: ";"
Private IMaxToken As Long  ' number of tokens
' Array of tokens after breaking up the string
Private Tokens() As String
'  Current character cursor position in the string, used when you move sequentially from one token to the next
Private IPos As Long
' Current length of the string
Private ILen As Long
Public Property Get Text() As Variant
' Return the whole string
   Text = strText
End Property
Public Property Let Text(ByVal vNewValue As Variant)
' Assign the Text string
   strText = vNewValue
   ILen = Len(strText)  ' Update length of string
   IPos = 1           ' Reset character cursor position
' Break up the string at delimiters. The default delimiter is "|"
   Tokenise
End Property
Public Property Get Delimiter() As Variant
' Return the delimiter
   Delimiter = strDelimiter
End Property
Public Function TokenAt(TNum) As String
' Return the Token whose arrayIndex is TNum.
' Tokens in the array are indexed in the range 1..IMaxToken
   If (TNum > 0) And (TNum <= IMaxToken) Then
      TokenAt = Tokens(TNum)
   Else
      TokenAt = ""
   End If
End Function
Public Property Let Delimiter(ByVal vNewValue As Variant)
'  Assign new delimiter value
   strDelimiter = vNewValue
' Break up the string at delimiters.
  Tokenise
End Property
Private Sub Tokenise()
' Break up the string at the delimiters into the tokens.
' Save the tokens in the array indexing in the range 1..IMaxToken
   Dim i
   i = 0: IPos = 1
   Do Until IPos > ILen
      i = i + 1
      ' Make room for one more token. Preserve the previous tokens
      ReDim Preserve Tokens(i)
      ' Obtain the next token from the string
      Tokens(i) = GetToken
   Loop
   ' Update total number of Tokens
   IMaxToken = i
   ' Reset character Cursor position
   IPos = 1
End Sub
Public Sub ReplaceToken(TNum, NewToken)
' Replace a Token with a new substring
   If (TNum > 0) And (TNum <= IMaxToken) Then
      ' Replace directly in the Token array
      Tokens(TNum) = NewToken
      '  Rebuild the internal text string
      ReconstructText
   End If
End Sub
Private Sub ReconstructText()
' Reconstruct the internal string from the tokens and the delimiter
   Dim i
   strText = ""  ' Initialise temporary string
   For i = 1 To IMaxToken
      ' Concatenate this token to strText
      strText = strText & Tokens(i)
      ' Append a delimiter character except for the last token
      If i < IMaxToken Then strText = strText & strDelimiter
   Next
End Sub
Public Function GetToken() As String
' Return the next token
' Use this function when you move sequentially from one token to the next
   Dim Pos
   GetToken = ""
   ' If the delimiter is a blank space, treat multiple blank spaces as a single
   '   blank space, i.e. skip the additional blank spaces
   If strDelimiter = " " Then
      Do While Mid(strText, IPos, 1) = " "
         IPos = IPos + 1
         ' Get out if reaching the end of the string
         If IPos > ILen Then
            Exit Function
         End If
      Loop
   End If
   ' Locate the next delimiter starting from the current character cursor position IPos
   Pos = InStr(IPos, strText, strDelimiter)
   If Pos > 0 Then
      ' extract the token substring from the string
      GetToken = Mid(strText, IPos, Pos - IPos)
      ' Update the character cursor position
      IPos = Pos + Len(strDelimiter)
   Else
      ' This is the last token so just use the remain of the string
      GetToken = Mid(strText, IPos, ILen - IPos + 1)
      ' Update the character cursor position
      IPos = ILen + 1
   End If
End Function
Public Function GetLastToken() As String
'Return the last token
   If (IMaxToken > 0) Then
      GetLastToken = TokenAt(IMaxToken)
   Else
      GetLastToken = ""
   End If
End Function
Public Property Get Length() As Long
' Return the length of the text string
   Length = ILen
End Property
Public Property Get TokenCount() As Variant
' Return the number of tokens
   TokenCount = IMaxToken
End Property

Private Sub Class_Initialize()
' The default delimiter is "|"
  strDelimiter = "|"
End Sub
