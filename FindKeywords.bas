Sub FindKeywords()
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim str As String
    Dim keyword As Variant
    
    ' Your input string
    str = "eager beaver:some are eager to meet me today. They won't be eager tomorrow, but grim in their faces." & vbCr _
          "grim ogre:meet me at the grim dine bar tonight. I love to enjoy good food." & vbCr _
          "enjoy:there are some people who lead a joyous life."

    ' List of keywords
    Dim keywords As Variant
    keywords = Array("grim", "eager", "enjoy", "some", "meet")

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")

    ' Loop through each keyword
    For Each keyword In keywords
        ' Set the pattern to search for the keyword after ":"
        regex.Pattern = "(?<=:" & vbCr & "|:).*\b" & keyword & "\b.*"
        
        ' Execute the regular expression
        Set matches = regex.Execute(str)
        
        ' Loop through the matches and print the results
        For Each match In matches
            MsgBox "Keyword: " & keyword & vbCr & "Matched Line: " & match.Value
        Next match
    Next keyword
End Sub
