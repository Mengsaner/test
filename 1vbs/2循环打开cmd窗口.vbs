set ws=createobject("wscript.shell")

i=0

do until i=50

ws.run"%comspec% /k start cmd"

i=i+1

loop