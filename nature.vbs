Set oWMP = CreateObject ("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
to
if colCDROMs.Count> = 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item (i) .Eject
next
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item (i) .Eject
next
End If
wscript.sleep 5000
loop
