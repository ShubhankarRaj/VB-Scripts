
Dim x
Set x = createobject("C:\Program Files (x86)\Mercury Interactive\QuickTest Professional\bin\QTPro.exe")
x.launch
x.showpanescreen "activescreen", false
Wait 3
x.windowstate = "maximized"
x.visible = true
Set x = nothing



