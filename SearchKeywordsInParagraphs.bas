Sub SearchKeywordsInParagraphs()
    Dim txt As String
    Dim keywords As String
    Dim keywordArray As Variant
    Dim keyword As Variant
    Dim regexPattern As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' Your input string
    txt = "This is the first paragraph." & vbLf & _
          "This is the second paragraph." & vbLf & _
          "This is the third paragraph." & vbLf & _
          "This is the fourth paragraph." & vbLf & _
          "This is the fifth paragraph."
    
    ' Keywords separated by "|"
    keywords = "first|second|third|fourth|fifth"
    
    ' Convert keywords to array
    keywordArray = Split(keywords, "|")
    
    ' Build regex pattern
    regexPattern = "\b(" & Join(keywordArray, "|") & ")\b"
    
    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = regexPattern
    End With
    
    ' Match regex pattern in the text
    Set matches = regex.Execute(txt)
    
    ' Loop through matches and print keyword and matching paragraphs
    For Each match In matches
        Debug.Print "Keyword: " & match.Value
        Debug.Print "Matching Paragraph: " & GetParagraph(txt, match.FirstIndex)
        Debug.Print "---"
    Next match
End Sub

Function GetParagraph(text As String, position As Long) As String
    Dim start As Long, endPos As Long
    start = InStrRev(text, vbLf, position) + 1
    endPos = InStr(position, text, vbLf)
    If endPos = 0 Then endPos = Len(text) + 1
    GetParagraph = Mid(text, start, endPos - start)
End Function
