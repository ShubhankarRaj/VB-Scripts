Set GoogBrowser = Description.Create
GoogBrowser("title").value = "Google"

Set GoogPage = Description.Create
GoogPage("title").value = "Google"

Set GoogWebEdit = Description.Create
GoogWebEdit("html tag").value = "INPUT"
GoogWebEdit("name").value = "q"

Browser(GoogBrowser).Page(GoogPage).WebEdit(GoogWebEdit).Set "DP is great"

'Using a wild character to click an image getAllAttributes.jpg
Browser(GoogBrowser).Page(GoogPage).Image("file name:=getAll.*").Click

Set oBrowser = Description.Create
oBrowser("title").value = "Google"

Set oPage = Description.Create
oPage()
