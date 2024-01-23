# Microsoft365R 2.4.0.9000

## Planner

- Fix a bug in the `ms_plan$get_details()` method.

# Microsoft365R 2.4.0

## OneDrive/SharePoint

- Fix broken functionality for shared items in OneDrive/Sharepoint. In particular, this should allow using the MS365 backend with the pins package (#149, #129).
- The `list_shared_items()`/`list_shared_files()` method for drives now always returns a list of drive item objects, rather than a data frame. If the `info` argument is supplied with a value other than "items", a warning is issued.
- Add folder upload and download functionality for `ms_drive_item$upload()` and `download()`. Subfolders can also be transferred recursively, and optionally in parallel. There are also corresponding `ms_drive$upload_folder()` and `download_folder()` methods.
- Add convenience methods for saving and loading datasets and R objects: `save_dataframe()`, `save_rds()`, `save_rdata()`, `load_dataframe()`, `load_rds()`, and `load_rdata()`. See `?ms_drive_item` and `?ms_drive` for more details.
- Add `copy` and `move` methods for drive items, and corresponding `copy_item` and `move_item` methods for drives.
- Add ability to upload and download via connections/raw vectors instead of files. You can specify the source to `ms_drive_item$upload()` to be a raw or text connection. Similarly, if the destination for `ms_drive_item$download()` is NULL, the downloaded data is returned as a raw vector.
- Add the ability to use object IDs instead of file/folder paths in `ms_drive` methods, including getting, uploading and downloading. This can be useful since the object ID is immutable, whereas file paths can change, eg if the file is moved or renamed. See `?ms_drive` for more details.

## Outlook

- Fix a bug in specifying multiple addresses in an email (#134, #151).
- Fix multiple inline images not showing up in emails (#107). Many thanks to @vorpalvorpal.

# Microsoft365R 2.3.4

- Fix a bug in retrieving a drive by name (#104).

# Microsoft365R 2.3.3

- Compatibility update for emayili version 0.6+. Note that this _breaks_ compatibility with emayili versions 0.5 and earlier.

# Microsoft365R 2.3.2

## OneDrive/SharePoint

- Add a `get_path()` method for drive items, which returns the path to the item starting from the root. Needed as Graph doesn't seem to store the path in an unmangled form anywhere.
- Fix broken methods for accessing items in shared OneDrive/SharePoint folders (#89).

## Teams

- Fix a bug in sending file attachments in Teams chats (#87).

## Other

- Add a vignette "Using Microsoft365R in an unattended script", describing the two options for scripting Microsoft365R: with a service principal, and with a service account.

# Microsoft365R 2.3.1

## OneDrive/SharePoint

- Add a `get_parent_folder()` method for drive items, which returns the parent folder as another drive item. The parent of the root is itself.

## Teams

- Fix a bug where attaching a file to a Teams chat/channel message would fail if the file was a type recognised by Microsoft 365 (#73).

## Other

- Add a vignette "Using Microsoft365R in a Shiny app" for this common use case.
- Make `token` an explicit argument to the client functions, for supplying an OAuth token object directly. Note that this was always possible, but is now better documented and supported. This is mostly to support the Shiny use case, as well as other situations where authentication is more complicated than usual.
- Changes to allow Microsoft365R to be usable without being on the search list (#72). Among other things, `make_basic_list()` is now a private method, rather than being exported. Thanks to Robert Ashton (@r-ash) for the PR.

# Microsoft365R 2.3.0

## Outlook

- Add support for shared mailboxes to `get_business_outlook()` (#39). To access a shared mailbox, supply one of the arguments `shared_mbox_id`, `shared_mbox_name` or `shared_mbox_email` specifying the ID, displayname or email address of the mailbox respectively.
- Fix a bug where the presence of calendar invites in an email folder caused `list_emails()` to fail (#60).

## Teams

- Add support for Teams chats (including one-on-one, group and meeting chats).
  - Use the `list_chats()` function to list the chats you're participating in, and the `get_chat()` function to retrieve a specific chat.
  - A chat object has class `ms_chat`, which has similar methods to a channel: you can send, list and retrieve messages, and list and retrieve members/attendees. One difference is that chats don't have an associated file folder, unlike channels.

# Microsoft365R 2.2.1

- Hotfix for `could not find function "make_basic_list"` error when calling `list_*` methods and functions (#58, #56).

# Microsoft365R 2.2.0

## OneDrive/SharePoint

- Add a `list_shared_items()` method for the `ms_drive` class to access files and folders shared with you (#45).
- Allow getting drives for groups, sites and teams by name. The first argument to the `get_drive()` method for these classes is now `drive_name`; to get a drive by ID, specify the argument name explicitly: `get_drive(drive_id=*)`
- Add a `by_item` argument to the `delete_item()` method for drives and the `delete()` method for drive items (#21). This is to allow deletion of non-empty folders on SharePoint sites with data protection policies in place. Use with caution.

## Outlook

- Add a `search` argument to the `ms_outlook_folder$list_emails()` method. The default is to search in the from, subject and body of the emails.

## Teams

- Add `list_members()` and `get_member()` methods for teams and channels.
- Add support for @mentions in Teams channel messages (#26).

## Other

- All `list_*` class methods now have `filter` and `n` arguments to filter the result set and cap the number of results, following the pattern in AzureGraph 1.3.0. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, an `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results. Note that support for filtering in the underlying Graph API is somewhat uneven at the moment.
- Experimental **read-only** support for plans, contributed by Roman Zenka.
  - Add `get_plan()` and `list_plans()` methods to the `az_group` class. Note that only Microsoft 365 groups can have  plans, not any other type of group.
  - To get the plan(s) for a site or team, call its `get_group()` method to retrieve the associated group, and then get the plan from the group.
  - A plan has methods to retrieve tasks and buckets, as well as plan details.

# Microsoft365R 2.1.0

- Add support for sending and managing emails in Outlook. Use the `get_personal_outlook()` and `get_business_outlook()` client functions to access the emails in your personal account and work or school account, respectively. Functionality supported includes:
  - Send and reply to emails, optionally composed with either the blastula or emayili packages
  - List and retrieve emails
  - Create and delete folders
  - Move and copy emails between folders
  - Move and copy folders
  - Add, remove, and download attachments
- Add ability to created nested folders in OneDrive and SharePoint document libraries (#24).
- Fix a bug that caused the `list_files()` method to fail on non-Windows systems (reported by Tony Sokolov).

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
  - `sharepoint_site()` is now `get_sharepoint_site()`
  - `personal_onedrive()` is now `get_personal_onedrive()`
  - `business_onedrive()` is now `get_business_onedrive()`
- The first argument to `get_sharepoint_site()` is `site_name` to get a site by name, for consistency with `get_team()`. To get a site by URL, specify the `site_url` argument explicitly: `get_sharepoint_site(site_url="https://my-site-url")`.
- Add `list_sharepoint_sites()` function to list the sites you follow.

## Other changes

- Add `bulk_import()` method for lists, for creating multiple items at once. Supply a data frame as the argument.
- The various client functions can now share the same underlying Graph login, which should reduce the incidence of token refreshing.

# Microsoft365R 1.0.0

- Initial CRAN release.
