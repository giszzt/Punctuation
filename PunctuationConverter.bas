' ====================================================
' Punctuation Converter for Word VBA
' ====================================================

Option Explicit

Sub ConvertPunctuation()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim txt As String
    Dim newTxt As String
    Dim changeCount As Long
    
    Set doc = ActiveDocument
    changeCount = 0
    
    Application.ScreenUpdating = False
    
    ' Process each paragraph
    For Each para In doc.Paragraphs
        Set rng = para.Range
        txt = rng.Text
        
        ' Skip empty paragraphs
        If Len(txt) > 1 Then
            newTxt = ConvertText(txt, changeCount)
            
            If newTxt <> txt Then
                ' Preserve the paragraph mark logic
                If Right(txt, 1) = vbCr Then
                    rng.Text = Left(newTxt, Len(newTxt) - 1) & vbCr
                Else
                    rng.Text = newTxt
                End If
            End If
        End If
    Next para
    
    Application.ScreenUpdating = True
    
    MsgBox "Converted " & changeCount & " punctuation marks.", vbInformation, "Done"
End Sub

Private Function ConvertText(ByVal txt As String, ByRef countRef As Long) As String
    Dim i As Long
    Dim char As String, prevChar As String, nextChar As String
    Dim result As String
    Dim newChar As String
    ' State tracking for quotes
    Dim inDouble As Boolean
    Dim inSingle As Boolean
    
    result = ""
    inDouble = False
    inSingle = False
    
    ' First pass: Detect initial state based on existing Chinese quotes
    ' This helps if the paragraph starts in the middle of a quote (rare) or has mixed content
    ' But for simplicity and robustness in standard writing, we usually assume False at paragraph start
    
    For i = 1 To Len(txt)
        char = Mid(txt, i, 1)
        prevChar = ""
        nextChar = ""
        
        If i > 1 Then prevChar = Mid(txt, i - 1, 1)
        If i < Len(txt) Then nextChar = Mid(txt, i + 1, 1)
        
        newChar = char
        
        ' 1. Update State based on existing Chinese quotes (Mixed content handling)
        If char = ChrW(&H201C) Then inDouble = True  ' “
        If char = ChrW(&H201D) Then inDouble = False ' ”
        If char = ChrW(&H2018) Then inSingle = True  ' ‘
        If char = ChrW(&H2019) Then inSingle = False ' ’
        
        ' 2. Check for ellipsis (...)
        If i <= Len(txt) - 2 Then
            If Mid(txt, i, 3) = "..." Then
                If IsChineseContext(prevChar) Or IsChineseContext(Mid(txt, i + 3, 1)) Then
                    result = result & ChrW(&H2026) & ChrW(&H2026)
                    countRef = countRef + 1
                    i = i + 2
                    GoTo NextChar
                End If
            End If
        End If
        
        ' 3. Check for dash (--)
        If i <= Len(txt) - 1 Then
            If Mid(txt, i, 2) = "--" Then
                If IsChineseContext(prevChar) Or IsChineseContext(Mid(txt, i + 2, 1)) Then
                    result = result & ChrW(&H2014) & ChrW(&H2014)
                    countRef = countRef + 1
                    i = i + 1
                    GoTo NextChar
                End If
            End If
        End If
        
        ' 4. Basic punctuation
        If IsChineseContext(prevChar) Or IsChineseContext(nextChar) Then
            Select Case char
                Case ","
                    newChar = ChrW(&HFF0C)
                    countRef = countRef + 1
                Case "."
                    newChar = ChrW(&H3002)
                    countRef = countRef + 1
                Case ":"
                    newChar = ChrW(&HFF1A)
                    countRef = countRef + 1
                Case ";"
                    newChar = ChrW(&HFF1B)
                    countRef = countRef + 1
                Case "?"
                    newChar = ChrW(&HFF1F)
                    countRef = countRef + 1
                Case "!"
                    newChar = ChrW(&HFF01)
                    countRef = countRef + 1
                Case "("
                    If IsChineseContext(nextChar) Then
                        newChar = ChrW(&HFF08)
                        countRef = countRef + 1
                    End If
                Case ")"
                    If IsChineseContext(prevChar) Then
                        newChar = ChrW(&HFF09)
                        countRef = countRef + 1
                    End If
                Case """"
                    ' Smart Toggle for Double Quotes
                    ' Logic: If we are not in a quote, open it. If we are, close it.
                    If Not inDouble Then
                        newChar = ChrW(&H201C) ' Open “
                        inDouble = True
                    Else
                        newChar = ChrW(&H201D) ' Close ”
                        inDouble = False
                    End If
                    countRef = countRef + 1
                Case "'"
                    ' Smart Toggle for Single Quotes
                    If Not inSingle Then
                        newChar = ChrW(&H2018) ' Open ‘
                        inSingle = True
                    Else
                        newChar = ChrW(&H2019) ' Close ’
                        inSingle = False
                    End If
                    countRef = countRef + 1
            End Select
        End If
        
        result = result & newChar
        
NextChar:
    Next i
    
    ConvertText = result
End Function

Private Function IsChineseContext(char As String) As Boolean
    Dim c As Long
    
    If Len(char) = 0 Then
        IsChineseContext = False
        Exit Function
    End If
    
    c = AscW(char) And &HFFFF&
    
    Select Case c
        Case &H4E00& To &H9FFF&
            IsChineseContext = True
        Case &H3000& To &H303F&
            IsChineseContext = True
        Case &HFF00& To &HEFFF&
            IsChineseContext = True
        Case Else
            IsChineseContext = False
    End Select
End Function
