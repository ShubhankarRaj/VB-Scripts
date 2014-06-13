'Dim regEx, Match, Matches
'cnt= 0
'stringPatt = "retsdgfggakfdakakdakdhadeepdakha"
'Set regEx = New RegExp
'regEx.Pattern = "d?ak?ha?"
'regEx.IgnoreCase = True
'regEx.Global = True
'
'Set Matches = regEx.Execute(stringPatt)
'Msgbox Matches.count
'For each i in Matches
'	result = result&i.value
'Next
'Msgbox result
'
Dim yourstr,r,letter,result
yourstr="Automation  45"
Set r=new regexp
r.pattern="[^abc]"
'r.pattern="a"
r.global=true
r.IgnoreCase = True
set s=r.execute(yourstr)
For each i in s
    result= i.value&result
Next 
msgbox result
















