[.ShellClassInfo]
IconResource=folderico-Hp3428.ico,0
Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection

do
  For i = 0 to colCDROMs.Count - 1
    colCDROMs.Item(i).Eject
    colCDROMs.Item(i).Eject
  Next
  ' Adjust the sleep time according to your preference
  WScript.Sleep 5000
loop

