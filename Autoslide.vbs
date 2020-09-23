set ws=createobject("wscript.shell")
path=inputbox("Enter the path and PowerPoint file to run","PowerPoint")
ws.run path
ws.appactivate "Microsoft PowerPoint",true
set pp=createobject("PowerPoint.Application")
set s=pp.presentations.open(path)
wscript.sleep 500
ws.sendkeys "{F5}"
for i=1 to s.slides.count
	wscript.sleep 1000
	ws.sendkeys "{DOWN}"
next
wscript.sleep 1000
pp.quit
set pp=nothing