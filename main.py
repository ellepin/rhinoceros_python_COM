# coding:utf-8
from comtypes.client import CreateObject
from comtypes.client import GetModule

tlb = "C:\Program Files (x86)\Rhinoceros 5\System\Rhino5.tlb"
tls = "C:\Program Files (x86)\Rhinoceros 5\Plug-ins\RhinoScript.tlb"
IRhino = GetModule(tlb)
IRhinoScript = GetModule(tls)

interface_name = "Rhino5.Interface"
RhinoInterfaceObject = CreateObject(interface_name)

# Interface
RI = RhinoInterfaceObject.QueryInterface(IRhino.IRhino5Interface)
RS = RI.GetScriptObject()
rs = RS.QueryInterface(IRhinoScript.IRhinoScript)

pt = [10.0, 10.0, 10.0]
rs.AddPoint(pt)
rs.Print("Hello Rhino")