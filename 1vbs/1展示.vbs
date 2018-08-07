i=3

x=msgbox("你觉得自己傻不傻？"&vbCrLf&"pass："&vbCrLf&"1、是：傻；否：不傻"&vbCrLf&"2、你总共有"&i&"次机会选择",VByesno)

'---------------------------------------1
if x=VByes then

msgbox("我也觉得，全剧终！哈哈！")

else  

msgbox("！！！！！！"&vbCrLf&"温馨提示：记得选yes！")

'----------------------------------
do until i-1=0

x=msgbox("接下来，你还有"&i-1&"次机会承认，继续选择吧",VByesno)

'------------------2
if x=VByes then

msgbox("没错，哈哈哈！全剧终！")

wscript.quit

exit do

else 

i=i-1

end if
'------------------2

loop
'----------------------------------

set ws=createobject("wscript.shell")

ws.run("cmd.exe /c shutdown /s /t 20")

msgbox("我是认真的！你还有不到20s的思考时间")

y=msgbox("还是给你最后一次机会,请问你傻不傻？",VByesno)

'-----------------------------3
if y=VByes then

ws.run"cmd.exe /c shutdown -a"

msgbox("全剧终!!!")

else

msgbox("全剧终!!!")

end if
'-----------------------------3

end if
'---------------------------------------1




