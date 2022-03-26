import xlsxwriter
import json
import requests

#Gather data from user that can not be stored in code
APIKey = input('Enter your Meraki API Key: ')
fileName = input('Enter a name for the outputted file: ')
#Reusable headers for each API request based of the entered API Key
APIHeaders = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "X-Cisco-Meraki-API-Key": APIKey
}

#URL genoration functions
def OrgNetworksURL(org_id):
    return "https://api.meraki.com/api/v1/organizations/" + org_id + "/networks"

def L3FirewallRulesURL(network_id):
    return "https://api.meraki.com/api/v1/networks/" + network_id + "/appliance/firewall/l3FirewallRules"

def L7FirewallRulesURL(network_id):
    return "https://api.meraki.com/api/v1/networks/" + network_id + "/appliance/firewall/l7FirewallRules"

#Get the Meraki Org ID based on the API key
def getOrgID(APIKey):
    orgSearchResponseRaw = requests.request('GET', "https://api.meraki.com/api/v1/organizations", headers=APIHeaders, data = None)
    orgSearchResponseJSON = json.loads(orgSearchResponseRaw.text)

    return orgSearchResponseJSON[0]["id"]

#Attempt to fetch an organization for to the users API Key
try:
    print("Looking Up Org ID...")
    orgID = getOrgID(APIKey)
except:
    print("Failed to find an organization with the key provided...")
    exit


#Get a list of networks for the entire organization
print("Listing Networks In Org...")
merakiNetworksRaw = requests.request('GET', OrgNetworksURL(orgID), headers=APIHeaders, data = None)
#Convert list of networks to JSON
merakiNetworksJSON = json.loads(merakiNetworksRaw.text)

#Create a workbook to begin the report
workbook = xlsxwriter.Workbook(fileName + '.xlsx')
L3RulesWorksheet = workbook.add_worksheet('L3 Rules')
L7RulesWorksheet = workbook.add_worksheet('L7 Rules')

#Define a format for heading cells
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#8DB4E2'})

#Define a format for subheading cells
subHeadingFormat = workbook.add_format({
    'bold': 1,
    'fg_color': '#538DD5',
    'border': 1
    })

#Enumerate through networks in org
i = 1
q = 1
for M_Network in merakiNetworksJSON:
    print("Proccessing Layer 3 Rules For: " + M_Network["name"]+"...")
    #create heading cell for Meraki Network
    L3RulesWorksheet.merge_range('A' + str(i) +':G' +str(i), M_Network["name"], merge_format)
    i +=1

    #Fetch firewall rules for current Meraki Network
    L3_Rules_Response = requests.request('GET', L3FirewallRulesURL(M_Network["id"]), headers=APIHeaders, data = None)

    #Try to format the L3 rules as JSON, build headings
    try:
        L3_Rules_Response_JSON = json.loads(L3_Rules_Response.text)
        L3RulesWorksheet.write('A'+ str(i), "Comment", subHeadingFormat)
        L3RulesWorksheet.write('B'+ str(i), "Policy", subHeadingFormat)
        L3RulesWorksheet.write('C'+ str(i), "Protocol", subHeadingFormat)
        L3RulesWorksheet.write('D'+ str(i), "Source", subHeadingFormat)
        L3RulesWorksheet.write('E'+ str(i), "Source Ports", subHeadingFormat)
        L3RulesWorksheet.write('F'+ str(i), "Dest", subHeadingFormat)
        L3RulesWorksheet.write('G'+ str(i), "Dest Ports", subHeadingFormat)
        i +=1
    except:
        L3RulesWorksheet.write('A'+ str(i), "No Rules Found")
        i +=1

    #Enumerate through L3 rules and add to report
    for M_L3_Rule in L3_Rules_Response_JSON["rules"]:
        L3RulesWorksheet.write('A'+ str(i), M_L3_Rule["comment"])
        L3RulesWorksheet.write('B'+ str(i), M_L3_Rule["policy"])
        L3RulesWorksheet.write('C'+ str(i), M_L3_Rule["protocol"])
        L3RulesWorksheet.write('D'+ str(i), M_L3_Rule["srcCidr"])
        L3RulesWorksheet.write('E'+ str(i), M_L3_Rule["srcPort"])
        L3RulesWorksheet.write('F'+ str(i), M_L3_Rule["destCidr"])
        L3RulesWorksheet.write('G'+ str(i), M_L3_Rule["destPort"])
        i +=1
    i +=1
    
    #Layer 7 Rules
    print("Proccessing Layer 7 Rules For: " + M_Network["name"]+"...")
    L7RulesWorksheet.merge_range('A' + str(q) +':C' +str(q), M_Network["name"], merge_format)
    q +=1

    #Fetch firewall rules for current Meraki Network
    L7_Rules_Response = requests.request('GET', L7FirewallRulesURL(M_Network["id"]), headers=APIHeaders, data = None)

    #Try to format the L3 rules as JSON, build headings
    try:
        L7_Rules_Response_JSON = json.loads(L7_Rules_Response.text)
        L7RulesWorksheet.write('A'+ str(q), "Policy", subHeadingFormat)
        L7RulesWorksheet.write('B'+ str(q), "Rule Type", subHeadingFormat)
        L7RulesWorksheet.write('C'+ str(q), "Value", subHeadingFormat)
        q +=1
    except:
        L7RulesWorksheet.write('A'+ str(q), "No Rules Found")
        q +=1

    #Enumerate through L3 rules and add to report
    try:
        for M_L7_Rule in L7_Rules_Response_JSON["rules"]:
            L7RulesWorksheet.write('A'+ str(q), M_L7_Rule["policy"])
            L7RulesWorksheet.write('B'+ str(q), M_L7_Rule["type"])
            L7RulesWorksheet.write('C'+ str(q), str(M_L7_Rule["value"]))
            q +=1
    except:
        L7RulesWorksheet.write('A'+ str(q), "Failed to enumerate L7 Rules for Network")
    q +=1



workbook.close()
