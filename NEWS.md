# Microsoft365R 1.0.0.9000

- Add support for Teams:
  - List channels
  - List messages and replies, send messages to channels
  - Upload and download files
- Add `bulk_import()` method for lists, for creating multiple items at once. Supply a data frame as the argument.
- Add `list_items()`/`list_files()` method for drive items, to list the files in a folder.
- The various client functions can now share the same underlying Graph login, which should reduce the incidence of token refreshing.

# Microsoft365R 1.0.0

- Initial CRAN release.
