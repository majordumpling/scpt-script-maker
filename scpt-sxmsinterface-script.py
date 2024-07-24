import pandas as pd

# Read XL file
df = pd.read_excel('C:\\Users\\jcarandang.CSCMWS\\Desktop\\SXMSINTERFACE Data.xlsx')

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

for index, row in df.iterrows():
    script = script + "\n\n'This code block adds entry " + row['INTERFACE']
    script = script + "\nsession.findById(\"wnd[0]\").maximize"
    script = script + "\nsession.findById(\"wnd[0]/tbar[1]/btn[5]\").press"
    script = script + "\nsession.findById(\"wnd[0]/usr/txtSXMSINTERFACE-INTERFACE\").text = \"" + row['INTERFACE'] + "\""
    if pd.notnull(row['LONGNAME']):
        script = script + "\nsession.findById(\"wnd[0]/usr/txtSXMSINTERFACE-LONGNAME\").text = \"" + str(row['LONGNAME']) + "\""
    if pd.notnull(row['PARTYAGENCY']):
        script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-PARTYAGENCY\").text = \"" + str(row['PARTYAGENCY']) + "\""
    if pd.notnull(row['PARTYTYPE']):
        script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-PARTYTYPE\").text = \"" + str(row['PARTYTYPE']) + "\""
    if pd.notnull(row['PARTY']):
        script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-PARTY\").text = \"" + str(row['PARTY']) + "\""
    script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-SERVICE\").text = \"" + row['SERVICE'] + "\""
    script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-NAME\").text = \"" + row['NAME'] + "\""
    script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-NAMESPACE\").text = \"" + row['NAMESPACE'] + "\""
    script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-SERVICE\").setFocus"
    script = script + "\nsession.findById(\"wnd[0]/usr/ctxtSXMSINTERFACE-SERVICE\").caretPosition = 4"
    script = script + "\nsession.findById(\"wnd[0]\").sendVKey 11"
    script = script + "\nsession.findById(\"wnd[0]\").sendVKey 3"
#with open('output.txt', 'w') as file:
#    file.write(repr(script))

print(script)
