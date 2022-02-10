dim a,s

set wshshell=createobject("wscript.shell")

a=0

do

wshshell.run"notepad"

a=a+1

if a>9999 then

exit do

end if