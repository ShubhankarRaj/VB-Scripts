Function RegExTest(patrn, strng)
   Dim regEx, Match, Matches, RetNum
   Set regEx = New RegExp
   regEx.Pattern = patrn
   regEx.IgnoreCase=TRUE
   regEx.Global = TRUE
   Set Matches = regEx.Execute(strng)
   RetNum = Matches.Count
   RegExTest = RetNum
End Function
Msgbox(RegExTest("a","aaaaaaaaahhhhhhhhhhhhsgggggggg"))
Msgbox(RegExTest("h","aaaaaaaaahhhhhhhhhhhhsgggggggg"))

'
'Function RegExpTest(patrn, strng)
'   Dim regEx, Match, Matches, RetNum  ' Create variable.
'   Set regEx = New RegExp   ' Create regular expression.
'   regEx.Pattern = patrn   ' Set pattern.
'   regEx.IgnoreCase = False   ' Set case insensitivity.
'   regEx.Global = True   ' Set global applicability.
'   Set Matches = regEx.Execute(strng)   ' Execute search.
'	RetNum = Matches.Count
''   For Each Match in Matches   ' Iterate Matches collection.
''      RetStr = RetStr & "Match found at position "
''      RetStr = RetStr & Match.FirstIndex & ". Match Value is  " 
''      RetStr = RetStr & Match.Value & "."
''   Next
'   RegExpTest = RetNum
'End Function
'MsgBox(RegExpTest("a", "aaabbbbcccccccddddaaabbbbsssss"))
'MsgBox(RegExpTest("b", "aaabbbbcccccccddddaaabbbbsssss"))














