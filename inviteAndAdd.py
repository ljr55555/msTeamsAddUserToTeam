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
strGuestEmailAddress = 'test@example.com'

postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}

r = requests.post(strGraphAuthURL, data=postData)

strJSONResponse = r.text
if len(strJSONResponse) > 5:
    jsonResponse = json.loads(strJSONResponse)

    strAccessToken = jsonResponse['access_token']

    postHeader = {'Content-Type': 'application/json', 'Authorization': 'Bearer ' + strAccessToken}
    getHeader = {"Authorization": "Bearer " + strAccessToken}

    # Initiate invitation to create guest acccount
    # https://docs.microsoft.com/en-us/graph/api/resources/invitation?view=graph-rest-1.0 for properties
    strInvitationContent = "{\"invitedUserEmailAddress\": \"{}\",\"inviteRedirectUrl\": \"https://teams.microsoft.com\",\"sendInvitationMessage\": \"true\"}".format(strGuestEmailAddress)
    jsonPostData = json.loads(strInvitationContent)

    res = requests.post("https://graph.microsoft.com/v1.0/invitations", headers=postHeader,data=json.dumps(jsonPostData))
    #print(res.text)

    jsonGuestInvitationResult = json.loads(res.text)
    jsonGuestID = jsonGuestInvitationResult.get("invitedUser")
    strGuestID = jsonGuestID.get("id")
    print("Guest invited - ID {}".format(strGuestID))
    time.sleep(30)      # Wait so guest account is found when we try to add it to the group

    # Get group ID for desired Teams group
    resGroupID = requests.get("https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and (displayname eq '{}')".format(strTeamName), headers=getHeader)
    #print(resGroupID.text)
    jsonGroupInfo = json.loads(resGroupID.text)
    strGroupInfoData = jsonGroupInfo.get('value')

    jsonGroupInfo = strGroupInfoData[0]
    strGroupID = jsonGroupInfo.get('id')
    print("Team with Guests has id {}".format(strGroupID )  )

    # Add newly created guest to Teams group
    if strGuestID and strGroupID:
        strInvitee = {"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/{}".format(strGuestID)}

        resInviteGuestToTeam = requests.post("https://graph.microsoft.com/v1.0/groups/{}/members/$ref".format(strGroupID), headers=postHeader,data=json.dumps(strInvitee))
        print(resInviteGuestToTeam.text)
        if resInviteGuestToTeam.status_code is 204:
            print("Successfully added guest {} to Team {}".format(strGuestEmailAddress,strTeamName))
        else:
            print("Return code is {} attempting to add guest {} to Team {}".format(resInviteGuestToTeam.status_code,strGuestEmailAddress,strTeamName))
    else:
        print("Missing important info ... bailing!")
