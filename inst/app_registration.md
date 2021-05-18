## App registration details

Microsoft365R is registered as an app in the `aicatr` AAD tenant. Depending on your organisation's security policy, you may have to get an admin to grant it access to your tenant.  The app ID is **d44a05d5-c6a5-4bbb-82d2-443123722380** and the default requested permissions are

- User.Read
- Files.ReadWrite.All
- Group.ReadWrite.All
- Directory.Read.All
- Mail.Send
- Mail.ReadWrite
- AllSites.Manage
- email, profile, openid, offline_access

These are delegated permissions, not application-level permissions, so a signed-in user is required.

Alternatively, if the environment variable `CLIMICROSOFT365_AADAPPID` is set, Microsoft365R will use its value as the app ID for authenticating to the Microsoft 365 Business services (SharePoint and OneDrive for Business). You can also specify the app ID as an argument when calling `get_personal_onedrive()`, `get_business_onedrive()` et al.

