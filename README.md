# SharePointR

A simple yet powerful interface to [SharePoint 365](https://www.microsoft.com/en-au/microsoft-365/sharepoint/collaboration) sites and [OneDrive](https://www.microsoft.com/en-au/microsoft-365/onedrive/online-cloud-storage), leveraging the facilities provided by the [AzureGraph](https://cran.r-project.org/package=AzureGraph) package. Both personal OneDrive and OneDrive for Business are supported.

## Examples

```r
## personal OneDrive
od <- personal_onedrive()

# list files and folders
od$list_items()
od$list_items("Documents")
od$download_file("Documents/myfile.doc")


## OneDrive for Business
odb <- business_onedrive("mycompany")

# same methods as for personal OneDrive
odb$list_items()
odb$upload_file("somedata.txt", "data/somedata.txt")


## SharePoint Online site
site <- sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name", tenant="mycompany")

site$list_drives()
site_drv <- site$get_drive()
site_drv$list_items()

lst <- site$get_list("my-list")
lst$list_items()
```
