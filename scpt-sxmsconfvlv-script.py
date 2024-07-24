import pandas as pd

# Read XL file
df = pd.read_excel('C:\\Users\\jcarandang.CSCMWS\\Desktop\\SXMSCONFVLV Data.xlsx')

script = "If Not IsObject(application) Then"
script = script + "\nSet SapGuiAuto  = GetObject(\"SAPGUI\")"
script = script + "\nSet application = SapGuiAuto.GetScriptingEngine"
script = script + "\nEnd If"
script = script + "\nIf Not IsObject(connection) Then"
script = script + "\nSet connection = application.Children(0)"
script = script + "\nEnd If"
script = script + "\nIf Not IsObject(session) Then"
script = script + "\nSet session    = connection.Children(0)"
script = script + "\nEnd If"
script = script + "\nIf IsObject(WScript) Then"
script = script + "\nWScript.ConnectObject session,     \"on\""
script = script + "\nWScript.ConnectObject application, \"on\""
script = script + "\nEnd If"

script = script + "\n\nsession.findById(\"wnd[0]\").maximize"
for index, row in df.iterrows():
    script = script + "\n\n'This code block adds entry " + str(row['Subparameter'])
    script = script + "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/cmbSXMSCONFVLV-AREA[1," + str(index) + "]\").key = \"RUNTIME\""
    script = script + "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-PARAM[2," + str(index) + "]\").text = \"" + str(row['Parameters']) + "\""
    if pd.notnull(row['Subparameter']):
        script = script + "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-SUBPARAM[3," + str(index) + "]\").text = \"" + str(row['Subparameter']) + "\""
    script = script + "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-VALUE[5," + str(index) + "]\").text = \"" + row['Value'] + "\""

script = script + "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/cmbSXMSCONFVLV-AREA[1,1]\").setFocus"
script = script + "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-VALUE[5,1]\").caretPosition = 12"
script = script + "\nsession.findById(\"wnd[0]\").sendVKey 11"
#with open('output.txt', 'w') as file:
#    file.write(repr(script))

print(script)
