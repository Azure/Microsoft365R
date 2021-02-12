# Microsoft365R 1.0.0.9000

## Major user-facing changes

- Add support for Teams:
  - Add `list_teams()` and `get_team()` client functions. You can get a team by team name or ID.
  - Get, list, create and delete channels
  - List messages and replies
  - Send messages to channels, send replies to messages
  - Upload and download files
  - In this version only Teams channels are supported; chats between individuals may come later.
- Move implementations for file and folder methods to the `ms_drive_item` class.
  - This facilitates managing files for Teams channels, which have associated folders in a shared document library (drive)
  - The existing methods for the `ms_drive` class now call down to the `ms_drive_item` methods, with appropriate arguments; their behaviour should be unchanged
- Rename the client functions to allow for listing teams and sites. The original clients are still available but deprecated, and will be removed in a future version of the package.
  - `get_sharepoint_site()` is now `get_sharepoint_site()`
  - `get_personal_onedrive()` is now `get_personal_onedrive()`
  - `get_business_onedrive()` is now `get_business_onedrive()`
- The first argument to `get_sharepoint_site()` is `site_name` to get a site by name, for consistency with `get_team()`. To get a site by URL, specify the `site_url` argument explicitly: `get_sharepoint_site(site_url="https://my-site-url")`.
- Add `list_sharepoint_sites()` function to list all sites you follow.

## Other changes

- Add `bulk_import()` method for lists, for creating multiple items at once. Supply a data frame as the argument.
- Add `list_items()`/`list_files()` method for drive items, to list the files in a folder.
- The various client functions can now share the same underlying Graph login, which should reduce the incidence of token refreshing.

# Microsoft365R 1.0.0

- Initial CRAN release.
