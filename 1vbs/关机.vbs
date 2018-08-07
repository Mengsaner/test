
x=msgbox("是否确定关机?",VByesno)

if x=VByes then

createobject("wscript.shell").run("%comspec% /c shutdown /s /t 3600")

end if

