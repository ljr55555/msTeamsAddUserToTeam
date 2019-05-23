import time
import datetime
import requests
from requests_toolbelt.utils import dump
import json
from config import  strClientID, strClientSecret, strGraphAuthURL

################################################################################
# Function definitions
################################################################################

################################################################################
# End of functions
################################################################################
strTeamName = 'Team With Guests'
strLogonUserUPN = 'azuretest@example.com'

postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}

r = requests.post(strGraphAuthURL, data=postData)

strJSONResponse = r.text
if len(strJSONResponse) > 5:
    jsonResponse = json.loads(strJSONResponse)

    strAccessToken = jsonResponse['access_token']

    postHeader = {'Content-Type': 'application/json', 'Authorization': 'Bearer ' + strAccessToken}
    getHeader = {"Authorization": "Bearer " + strAccessToken}

    # Find the object ID for the logged on user
    strUserSearchURI = "https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '{}'".format(strLogonUserUPN)
    res = requests.get(strUserSearchURI, headers=getHeader)

    jsonFoundUser = json.loads(res.text)
    jsonFoundUserDetails = (jsonFoundUser.get("value"))[0]
    strFoundUserID = jsonFoundUserDetails.get("id")
    print("User to add has ID {}".format(strFoundUserID))

    # Get group ID for desired Teams group
    resGroupID = requests.get("https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and (displayname eq '{}')".format(strTeamName), headers=getHeader)
    #print(resGroupID.text)
    jsonGroupInfo = json.loads(resGroupID.text)
    strGroupInfoData = jsonGroupInfo.get('value')

    jsonGroupInfo = strGroupInfoData[0]
    strGroupID = jsonGroupInfo.get('id')
    print("Team with Guests has id {}".format(strGroupID )  )

    # Add newly created guest to Teams group
    if strFoundUserID and strGroupID:
        strInvitee = {"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/{}".format(strFoundUserID)}

        resInviteGuestToTeam = requests.post("https://graph.microsoft.com/v1.0/groups/{}/members/$ref".format(strGroupID), headers=postHeader,data=json.dumps(strInvitee))
        print(resInviteGuestToTeam.text)
        if resInviteGuestToTeam.status_code is 204:
            print("Successfully added guest {} to Team {}".format(strLogonUserUPN,strTeamName))
        else:
            strErrorText = json.loads(resInviteGuestToTeam.text)
            print(strErrorText)
            strErrorMessage = (strErrorText.get("error")).get("message")
            print("Error  is {} attempting to add guest {} to Team {}".format(strErrorMessage,strLogonUserUPN,strTeamName))
    else:
        print("Missing important info ... bailing!")
