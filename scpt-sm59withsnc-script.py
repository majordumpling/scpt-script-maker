#Imports
import pandas as pd
import importlib

# Initial variables declaration
module_name = "scpt-custom-module"
mod = importlib.import_module(module_name)

script = "If Not IsObject(application) Then"
script += "\nSet SapGuiAuto = GetObject(\"SAPGUI\")"
script += "\nSet application = SapGuiAuto.GetScriptingEngine"
script += "\nEnd If"
script += "\nIf Not IsObject(connection) Then"
script += "\nSet connection = application.Children(0)"
script += "\nEnd If"
script += "\nIf Not IsObject(session) Then"
script += "\nSet session = connection.Children(0)"
script += "\nEnd If"
script += "\nIf IsObject(WScript) Then"
script += "\nWScript.ConnectObject session,     \"on\""
script += "\nWScript.ConnectObject application, \"on\""
script += "\nEnd If"
script += "\n\nsession.findById(\"wnd[0]\").maximize"

#Function declarations
def createTypeT(local, rfcName, desc1, desc2, desc3, progId, gatewayHost, gatewayService, snc):
    local = "\n\n'This code block creates rfc " + rfcName
    local += "\nsession.findById(\"wnd[0]/usr/cntlSM59CNTL_AREA/shellcont/shell/shellcont[1]/shell[0]\").pressButton \"CREATE\""
    local += "\nsession.findById(\"wnd[1]/usr/txtDATA_010-DESTINATION_NAME\").text = \"" + rfcName + "\""
    local += "\nsession.findById(\"wnd[1]/usr/cmbDATA_010-DESTINATION_TYPE\").key = \"T\""
    local += "\nsession.findById(\"wnd[1]/usr/cmbDATA_010-DESTINATION_TYPE\").setFocus"
    local += "\nsession.findById(\"wnd[1]/tbar[0]/btn[0]\").press"
    local += "\nsession.findById(\"wnd[0]/usr/txtRFCDOC-RFCDOC1\").text = \"" + desc1 + "\""
    local += "\nsession.findById(\"wnd[0]/usr/txtRFCDOC-RFCDOC2\").text = \"" + desc2 + "\""
    local += "\nsession.findById(\"wnd[0]/usr/txtRFCDOC-RFCDOC3\").text = \"" + desc3 + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0510/radGL_TYPT_4\").setFocus"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0510/radGL_TYPT_4\").select"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0510/subSUB_TYPT:SAPLCRFC:0514/txtRFCREGID\").text = \"" + progId + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0510/ssubSUB_GW:SAPLCRFC:0520/txtRFCDISPLAY-RFCGWHOST\").text = \"" + gatewayHost + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0510/ssubSUB_GW:SAPLCRFC:0520/txtRFCDISPLAY-RFCGWSERV\").text = \"" + gatewayService + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0510/ssubSUB_GW:SAPLCRFC:0520/txtRFCDISPLAY-RFCGWSERV\").setFocus"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0510/ssubSUB_GW:SAPLCRFC:0520/txtRFCDISPLAY-RFCGWSERV\").caretPosition = 14"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN\").select"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0600/btnSNC\").press"
    local += "\nsession.findById(\"wnd[1]/usr/chkRFCDISPLAY-RFCSNC\").setFocus"
    local += "\nsession.findById(\"wnd[1]\").close"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0600/radRFCDISPLAY-RFCSNC\").setFocus"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0600/radRFCDISPLAY-RFCSNC\").select"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0600/btnSNC\").press"
    local += "\nsession.findById(\"wnd[1]/usr/ctxtRFCDESSECU-SNC_QOP\").text = \"8\""
    local += "\nsession.findById(\"wnd[1]/usr/txtRFCDESSECU-PNAME_APPL\").text = \"snc\""
    local += "\nsession.findById(\"wnd[1]/usr/txtRFCDESSECU-PNAME_APPL\").setFocus"
    local += "\nsession.findById(\"wnd[1]/usr/txtRFCDESSECU-PNAME_APPL\").caretPosition = 3"
    local += "\nsession.findById(\"wnd[1]/tbar[0]/btn[11]\").press"
    local += "\nsession.findById(\"wnd[2]/tbar[0]/btn[5]\").press"
    local += "\nsession.findById(\"wnd[1]/usr/txtRFCDESSECU-PNAME_APPL\").text = \"" + snc + "\""
    local += "\nsession.findById(\"wnd[1]/usr/txtRFCDESSECU-PNAME_APPL\").caretPosition = 25"
    local += "\nsession.findById(\"wnd[1]/tbar[0]/btn[11]\").press"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0600/radRFCDISPLAY-RFCSNC\").setFocus"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0600/radRFCDISPLAY-RFCSNC\").select"
    local += "\nsession.findById(\"wnd[0]/tbar[0]/btn[11]\").press"
    local += "\nsession.findById(\"wnd[0]/tbar[0]/btn[3]\").press"
    return local

def createTypeH(local, rfcName, desc1, desc2, desc3, host, port, prefix, password):
    local = "\n\n'This code block creates rfc " + rfcName
    local += "\nsession.findById(\"wnd[0]/usr/cntlSM59CNTL_AREA/shellcont/shell/shellcont[1]/shell[0]\").pressButton \"CREATE\""
    local += "\nsession.findById(\"wnd[1]/usr/txtDATA_010-DESTINATION_NAME\").text = \"" + rfcName + "\""
    local += "\nsession.findById(\"wnd[1]/usr/cmbDATA_010-DESTINATION_TYPE\").key = \"H\""
    local += "\nsession.findById(\"wnd[1]/usr/cmbDATA_010-DESTINATION_TYPE\").setFocus"
    local += "\nsession.findById(\"wnd[1]/tbar[0]/btn[0]\").press"
    local += "\nsession.findById(\"wnd[0]/usr/txtRFCDOC-RFCDOC1\").text = \"" + desc1 + "\""
    local += "\nsession.findById(\"wnd[0]/usr/txtRFCDOC-RFCDOC2\").text = \"" + desc2 + "\""
    local += "\nsession.findById(\"wnd[0]/usr/txtRFCDOC-RFCDOC3\").text = \"" + desc3 + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0530/txtHOSTNAME\").text = \"" + host + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0530/txtRFCDISPLAY-RFCSYSID\").text = \"" + port + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0530/txtRFCDISPLAY-PFADPRE\").text = \"" + prefix + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0530/txtRFCDISPLAY-PFADPRE\").setFocus"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0530/txtRFCDISPLAY-PFADPRE\").caretPosition = 6"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0530/txtRFCDISPLAY-PFADPRE\").text = \"" + prefix + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpTECH/ssubSUB_SM59:SAPLCRFC:0530/txtRFCDISPLAY-PFADPRE\").caretPosition = 7"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN\").select"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/radRFCDISPLAZ-RFCSLOGIN3\").setFocus"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/radRFCDISPLAZ-RFCSLOGIN3\").select"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/txtRFCDISPLAY-RFCLANG\").text = \"EN\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/txtRFCDISPLAY-RFCCLIENT\").text = \"400\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/txtRFCDISPLAY-RFCUSER\").text = \"SAPPI0A6A\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/pwdRFCPASSWORD\").text = \"" + password + "\""
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/radRFCDISPLAY-RFCSNC\").setFocus"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/radRFCDISPLAY-RFCSNC\").select"
    local += "\nsession.findById(\"wnd[1]/tbar[0]/btn[0]\").press"
    local += "\nsession.findById(\"wnd[0]/usr/tabsTAB_SM59/tabpSIGN/ssubSUB_SM59:SAPLCRFC:0610/cmbRFCDISPLAY-SSLAPPLIC\").setFocus"
    local += "\nsession.findById(\"wnd[0]/tbar[0]/btn[11]\").press"
    local += "\nsession.findById(\"wnd[0]/tbar[0]/btn[3]\").press"
    return local

def removeNan(str):
    if pd.isna(str):
        str = ""
    return str

#Main

# Call the function from the imported module
filepath = mod.get_valid_file()

df = pd.read_excel(filepath)

#Iterates thru each value and builds SAP script
for index, row in df.iterrows():
    if str(row['Type']) == "T":
        script += createTypeT(script, str(row['RFC Destination']), removeNan(str(row['RFCDOC1'])), removeNan(str(row['RFCDOC2'])), removeNan(str(row['RFCDOC3'])), str(row['Program ID']), str(row['Gateway host']), str(row['Gateway service']), str(row['SNC']))
    else:
        script += createTypeH(script, str(row['RFC Destination']), removeNan(str(row['RFCDOC1'])), removeNan(str(row['RFCDOC2'])), removeNan(str(row['RFCDOC3'])), str(row['Target host']), str(row['Port']), str(row['Path prefix']), str(row['PW']))

mod.save_file(script)