Attribute VB_Name = "mdlMain"
'Author: Le Viet Man - TinK24B - DHKH Hue
'Email: manleviet@yahoo.com
Option Explicit

Public Type THuffman
    ID As Integer
    HCode As String '* 40
    R As Integer
    L As Integer
    B As Integer
    p As Single
End Type

Public Huffman(1 To 32, 1 To 33) As THuffman
Public CharVn As Variant

Sub Main()
    CharVn = Array("a", "aw", "aa", "b", "c", "d", "dd", "e", "ee" _
    , "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "ow", "oo" _
    , "p", "q", "r", "s", "t", "u", "uw", "v", "x", "y", "w", "z")
    frmHuffman.Show
End Sub

Public Sub ReadInfor(H() As THuffman)
Dim i As Integer
    For i = 1 To 33
        H(1, i).ID = i
        H(1, i).p = CSng(frmHuffman.txtp(i - 1).Text)
    Next i
End Sub

Public Sub WriteInfor(H() As THuffman)
Dim i As Integer
Dim t As Single
    SortByID H, 1
    For i = 1 To 33
        t = t + (H(1, i).p * Len(H(1, i).HCode))
        frmHuffman.lblhuffman(i - 1).Caption = H(1, i).HCode
    Next i
    frmHuffman.lblhuffman1.Caption = t
End Sub

Public Sub SwapH(a As THuffman, c As THuffman)
Dim t As THuffman
     t = a
     a = c
     c = t
End Sub

Public Sub SortByp(H() As THuffman, t As Integer)
Dim i As Integer, j As Integer
    For i = 1 To 33 - t
        For j = i + 1 To 34 - t
            If H(t, i).p < H(t, j).p Then
                SwapH H(t, i), H(t, j)
            End If
        Next j
    Next i
End Sub

Public Sub FixSource(H() As THuffman)
Dim i As Integer, j As Integer
    SortByp H, 1
    For i = 1 To 31
        For j = 1 To 32 - i
            H(i + 1, j).ID = j
            H(i + 1, j).B = j
            H(i + 1, j).p = H(i, j).p
        Next j
        H(i + 1, j).ID = j
        H(i + 1, j).R = j
        H(i + 1, j).L = j + 1
        H(i + 1, j).p = H(i, j).p + H(i, j + 1).p
        SortByp H, i + 1
    Next i
End Sub

Public Sub HCoding(H() As THuffman)
Dim i As Integer, j As Integer
    H(32, 1).HCode = "0"
    H(32, 2).HCode = "1"
    For i = 32 To 2 Step -1
        For j = 1 To 34 - i
            If H(i, j).B <> 0 Then
                H(i - 1, Huffman(i, j).B).HCode = H(i, j).HCode
            Else
                H(i - 1, Huffman(i, j).R).HCode = H(i, j).HCode & "0"
                H(i - 1, Huffman(i, j).L).HCode = H(i, j).HCode & "1"
            End If
        Next j
    Next i
End Sub

Public Sub SortByID(Huffman() As THuffman, t As Integer)
Dim i As Integer, j As Integer
    For i = 1 To 33 - t
        For j = i + 1 To 34 - t
            If Huffman(t, i).ID > Huffman(t, j).ID Then
                SwapH Huffman(t, i), Huffman(t, j)
            End If
        Next j
    Next i
End Sub

Public Sub ClearH(H() As THuffman)
Dim i As Integer, j As Integer
    For i = 1 To 32
        For j = 1 To 33
            H(i, j).ID = 0
            H(i, j).HCode = ""
            H(i, j).p = 0
            H(i, j).R = 0
            H(i, j).L = 0
            H(i, j).B = 0
        Next j
    Next i
End Sub

Public Sub LoadInfor(cDialog As CommonDialog)
Dim fileNo, myLocalFolder, myFileName
Dim st As String
Dim DString As clsString
Dim i As Integer
    fileNo = FreeFile
    myFileName = cDialog.FileName
    Open myFileName For Input As #fileNo
    i = 0
    Do While Not EOF(fileNo)
        Line Input #fileNo, st
        Set DString = New clsString
        DString.Text = st
        DString.Delimiter = " "
        frmHuffman.txtp(i) = DString.TokenAt(3)
        i = i + 1
    Loop
    Close #fileNo
End Sub

Public Sub SaveInp(cDialog As CommonDialog)
Dim fileNo, myLocalFolder, myFileName
Dim i As Integer
    fileNo = FreeFile
    myFileName = cDialog.FileName
    Open myFileName For Output As #fileNo
    i = 0
    Do While i < 33
        Print #fileNo, i & " " & CharVn(i) & " " & frmHuffman.txtp(i).Text
        i = i + 1
    Loop
    Close #fileNo
End Sub

Public Sub SaveOut(cDialog As CommonDialog)
Dim fileNo, myLocalFolder, myFileName
Dim i As Integer
    fileNo = FreeFile
    myFileName = cDialog.FileName
    MsgBox myFileName
    Open myFileName For Output As #fileNo
    i = 0
    Do While i < 33
        Print #fileNo, i & " " & CharVn(i) & " " & frmHuffman.lblhuffman(i).Caption
        i = i + 1
    Loop
    Close #fileNo
End Sub

