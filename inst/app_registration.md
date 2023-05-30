## App registration details

Microsoft365R comes with a default app registration to enable authenticating with Azure Active Directory on the local machine. Depending on your organisation's security policy, you may have to get an admin to grant it access to your tenant.  The app ID is **d44a05d5-c6a5-4bbb-82d2-443123722380** and the default requested permissions are

- User.Read
- Files.ReadWrite.All
- Group.ReadWrite.All
- Directory.Read.All
- Mail.Send
- Mail.ReadWrite
- Sites.Manage.All
- Sites.ReadWrite.All
- email, profile, openid, offline_access

In addition, some functions request the following permissions:

- Mail.Send.Shared
- Mail.ReadWrite.Shared
- Chat.ReadWrite

These are Microsoft Graph permissions (`https://graph.microsoft.com/`). They are all delegated permissions, not application-level permissions, so a signed-in user is required. As the [Graph documentation](https://learn.microsoft.com/en-us/graph/auth/auth-concepts#microsoft-graph-permissions) notes:

> For delegated permissions, the effective permissions of your app are the intersection of the delegated permissions the app has been granted (via consent) and the privileges of the currently signed-in user. Your app can never have more privileges than the signed-in user. Within organizations, the privileges of the signed-in user are determined by policy or by membership in one or more administrator roles.

Note that the default app registration is only for a local machine: if you are running Microsoft365R inside a Shiny app on a remote server, it will not work. To enable interactive authentication, you must register your app with Azure and supply details such as the target audience and site address (redirect URI): see the Shiny vignette for more information.

