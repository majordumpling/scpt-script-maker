#Imports
import pandas as pd
import importlib

# Initial variables declaration
module_name = "scpt-custom-module"
mod = importlib.import_module(module_name)

script = "If Not IsObject(application) Then"
script += "\nSet SapGuiAuto  = GetObject(\"SAPGUI\")"
script += "\nSet application = SapGuiAuto.GetScriptingEngine"
script += "\nEnd If"
script += "\nIf Not IsObject(connection) Then"
script += "\nSet connection = application.Children(0)"
script += "\nEnd If"
script += "\nIf Not IsObject(session) Then"
script += "\nSet session    = connection.Children(0)"
script += "\nEnd If"
script += "\nIf IsObject(WScript) Then"
script += "\nWScript.ConnectObject session,     \"on\""
script += "\nWScript.ConnectObject application, \"on\""
script += "\nEnd If"
script += "\n\nsession.findById(\"wnd[0]\").maximize"

#Function declarations

#Main
#Asks user for location of excel file
filepath = mod.get_valid_file()

#Reads excel file and stores in a data frame variable
df = pd.read_excel(filepath)

#Iterates thru each value and builds SAP script
for index, row in df.iterrows():
    script += "\n\n'This code block adds entry " + str(row['Subparameter'])
    script += "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/cmbSXMSCONFVLV-AREA[1," + str(index) + "]\").key = \"RUNTIME\""
    script += "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-PARAM[2," + str(index) + "]\").text = \"" + str(row['Parameters']) + "\""
    if pd.notnull(row['Subparameter']):
        script += "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-SUBPARAM[3," + str(index) + "]\").text = \"" + str(row['Subparameter']) + "\""
    script += "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-VALUE[5," + str(index) + "]\").text = \"" + row['Value'] + "\""

script += "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/cmbSXMSCONFVLV-AREA[1,1]\").setFocus"
script += "\nsession.findById(\"wnd[0]/usr/tblSAPLSXMSCONFUITCTRL_SXMSCONFVLV/ctxtSXMSCONFVLV-VALUE[5,1]\").caretPosition = 12"
script += "\nsession.findById(\"wnd[0]\").sendVKey 11"

mod.save_file(script)
