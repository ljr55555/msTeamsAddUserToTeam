# msTeamsAddUserToTeam
Python scripts to add an existing tenant account to a Team (findAndAdd.py) and invite a guest to the tenant and add the newly created account to a Team (inviteAndAdd.py)

Requires an application registered in AD tenant. App API permissions -- Microsoft Graph:
* Directory.ReadWrite.All
* Group.ReadWrite.All
* User.Invite.All
* User.Read
* User.ReadWrite.All
