import pandas as pd
import importlib

# Read XL file
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

#Main
#Asks user for location of excel file
filepath = mod.get_valid_file()

#Reads excel file and stores in a data frame variable
df = pd.read_excel(filepath)

for index, row in df.iterrows():
    script += "\n\n'This code block adds entry " + row['INTERFACE']
    script += "\nsession.findById(\"wnd[0]\").maximize"
    script += "\nsession.findById(\"wnd[0]/tbar[1]/btn[5]\").press"
    script += "\nsession.findById(\"wnd[0]/usr/txtSXMSINTERFACE-INTERFACE\").text = \"" + row['INTERFACE'] + "\""
    if pd.notnull(row['LONGNAME']):
        script += "\nsession.findById(\"wnd[0]/usr/txtSXMSINTERFACE-LONGNAME\").text = \"" + str(row['LONGNAME']) + "\""
    if pd.notnull(row['PARTYAGENCY']):
        script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-PARTYAGENCY\").text = \"" + str(row['PARTYAGENCY']) + "\""
    if pd.notnull(row['PARTYTYPE']):
        script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-PARTYTYPE\").text = \"" + str(row['PARTYTYPE']) + "\""
    if pd.notnull(row['PARTY']):
        script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-PARTY\").text = \"" + str(row['PARTY']) + "\""
    script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-SERVICE\").text = \"" + row['SERVICE'] + "\""
    script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-NAME\").text = \"" + row['NAME'] + "\""
    script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-NAMESPACE\").text = \"" + row['NAMESPACE'] + "\""
    script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-SERVICE\").setFocus"
    script += "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-SERVICE\").caretPosition = 4"
    script += "\nsession.findById(\"wnd[0]\").sendVKey 11"
    script += "\nsession.findById(\"wnd[0]\").sendVKey 3"

mod.save_file(script)