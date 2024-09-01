import clr
clr.AddReference("./TwinCAT.Ads")
from TwinCAT.Ads import AmsAddress

obj = AmsAddress()

print(obj.ToString())
