# Microsoft365R 2.0.0.9000

- Add support for sending and managing emails in Outlook. Use the `get_personal_outlook()` and `get_business_outlook()` client functions to access the emails in your personal account and work or school account, respectively. Functionality supported includes:
  - Send and reply to emails, optionally composed with either the blastula or emayili packages
  - List and retrieve emails
  - Create and delete folders
  - Move and copy emails between folders
  - Move and copy folders
  - Add, remove, and download attachments
- Add ability to created nested folders in OneDrive and SharePoint document libraries (#24).
- Add a `by_item` argument to the `delete_item()` method for drives and the `delete()` method for drive items. This is to allow deletion of non-empty folders on SharePoint sites with data protection policies in place. Use with caution (#21).

# Microsoft365R 2.0.0

## Major user-facing changes

- Add `list_teams()` and `get_team()` client functions for working with Microsoft Teams. You can get a team by name or ID. The following Teams functionality is supported:
  - Get, list, create and delete channels
  - List messages and replies
  - Send messages to channels, send replies to messages
  - Upload and download files
  - In this version only Teams channels are supported; chats between individuals may come later.
- Move implementations for file and folder methods to the `ms_drive_item` class.
  - This includes the following: `list_files/list_items()`, `get_item()`, `create_folder()`, `upload()` and `download()`.
  - This facilitates managing files for Teams channels, which have associated folders in a shared document library (drive)
  - The existing methods for the `ms_drive` class now call down to the `ms_drive_item` methods, with appropriate arguments; their behaviour should be unchanged
- Rename the client functions to allow for listing teams and sites. The original clients are still available, but are deprecated and simply redirect to the new functions. They will be removed in a future version of the package.
  - `get_sharepoint_site()` is now `get_sharepoint_site()`
  - `get_personal_onedrive()` is now `get_personal_onedrive()`
  - `get_business_onedrive()` is now `get_business_onedrive()`
- The first argument to `get_sharepoint_site()` is `site_name` to get a site by name, for consistency with `get_team()`. To get a site by URL, specify the `site_url` argument explicitly: `get_sharepoint_site(site_url="https://my-site-url")`.
- Add `list_sharepoint_sites()` function to list the sites you follow.

## Other changes

- Add `bulk_import()` method for lists, for creating multiple items at once. Supply a data frame as the argument.
- The various client functions can now share the same underlying Graph login, which should reduce the incidence of token refreshing.

# Microsoft365R 1.0.0

- Initial CRAN release.
