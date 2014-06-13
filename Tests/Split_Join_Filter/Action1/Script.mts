

    Dim strInputString '     As String

    Dim strFilterText'       As String

    Dim astrSplitItems'    As String

    Dim astrFilteredItems' As String

    Dim strFilteredString'   As String

    Dim intX'                As Integer

   

    strInputString = InputBox("Enter a comma-delimited string of items: String Array Functions")

    strFilterText = InputBox("Enter Filter:", "String Array Functions")

   

    Print "Original Input String: "&strInputString

    Print vbNewLine

'	Dim MyString, MyArray, Msg
'	MyString = "VBScript X is X fun!"
'	MyArray = Split(MyString, "x", -1, 1)
'
'	Msg = MyArray(0) & " " & MyArray(1)
'	Msg = Msg   & " " & MyArray(2)
'	MsgBox Msg

    Print "Split Items:"

    astrSplitItems = Split(strInputString,",", -1,1)
	'split(

    For intX = 0 To UBound(astrSplitItems)

        Print "Item("&intX&"): "&astrSplitItems(intX)

    Next

    Print vbNewLine

   Print "Filtered Items (using '"&strFilterText&"'):"

   astrFilteredItems = Filter(astrSplitItems, strFilterText, True, vbTextCompare)

    For intX = 0 To UBound(astrFilteredItems)

        Print "Item("&intX&"): "&astrFilteredItems(intX)

    Next

    strFilteredString = Join(astrFilteredItems, ",")

    Print "Filtered Output String: "&strFilteredString