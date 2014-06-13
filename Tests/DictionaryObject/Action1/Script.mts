
Dim d   ' Create some variables.
   Set d = CreateObject("Scripting.Dictionary")
   d.CompareMode = vbTextCompare
   d.Add "z", "Athens"   ' Add some keys and items.
   d.Add "d", "Belgrade"
   d.Add "x", "Cairo"
   d.Key("x") = "h"   ' Set key for "c" to "d".
   DicDemo = d.Item("h")
   Msgbox DicDemo
   MsgBox d.Item("x")

   d.Key("h") = "f"   ' Set key for "c" to "d".
   DicDemo = d.Item("f")
   Msgbox DicDemo
   DicDemo = d.Item("h")
   MsgBox DicDemo
'
'	d.Remove "c"
'd.Key("e") = "c"   ' Set key for "c" to "d".
'   DicDemo = d.Item("c")
'   Msgbox DicDemo
'   MsgBox d.Item("e")
'	MsgBox d.Item("d")
'Dim a,b
'Dim oNAL : Set oNAL = CreateObject( "System.Collections.ArrayList" )
a = d.Items
b = d.keys
'For i = 0 to d.Count - 1
'	Msgbox a(i)
'    Msgbox b(i)
'	oNAL.Add b(i)
'	
'Next
'oNAL.Sort
'For i = 0 to d.Count - 1
'	Msgbox oNAL(i)
'    Msgbox d.Item(oNAL(i))
'Next

For i = 0 to d.count -1
	Msgbox a(i)&"--"&b(i)
Next