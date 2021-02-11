# Microsoft365R <img src="man/figures/logo.png" align="right" width=150 />

[![CRAN](https://www.r-pkg.org/badges/version/Microsoft365R)](https://cran.r-project.org/package=Microsoft365R)
![Downloads](https://cranlogs.r-pkg.org/badges/Microsoft365R)
![R-CMD-check](https://github.com/Azure/Microsoft365R/workflows/R-CMD-check/badge.svg)

Microsoft365R is intended to be a simple yet powerful R interface to [Microsoft 365](https://www.microsoft.com/en-us/microsoft-365) (formerly known as Office 365), leveraging the facilities provided by the [AzureGraph](https://cran.r-project.org/package=AzureGraph) package. Currently it enables access to data stored in [SharePoint Online](https://www.microsoft.com/en-au/microsoft-365/sharepoint/collaboration) sites and [OneDrive](https://www.microsoft.com/en-au/microsoft-365/onedrive/online-cloud-storage). Both personal OneDrive and OneDrive for Business are supported. Future versions may add support for Teams, Outlook and other Microsoft 365 services.

The primary repo for this package is at https://github.com/Azure/Microsoft365R; please submit issues and PRs there. It is also mirrored at the Cloudyr org at https://github.com/cloudyr/Microsoft365R. You can install the development version of the package with `devtools::install_github("Azure/Microsoft365R")`.

## Authentication details

The first time you call one of the Microsoft365R functions (see below), it will use your Internet browser to authenticate with Azure Active Directory, in a similar manner to other web apps. See [app_registration.md](https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md) for more details on the app registration and permissions requested.

## OneDrive

To access your personal OneDrive, call the `get_personal_onedrive()` function. This returns an R6 client object of class `ms_drive`, which has methods for working with files and folders.

```r
od <- get_personal_onedrive()

# list files and folders
od$list_items()
od$list_items("Documents")

# upload and download files
od$download_file("Documents/myfile.docx")
od$upload_file("somedata.xlsx")

# create a folder
od$create_folder("Documents/newfolder")
```

You can open a file or folder in your browser with the `open_item()` method. For example, a Word document or Excel spreadsheet will open in Word or Excel Online, and a folder will be shown in OneDrive.

```r
od$open_item("Documents/myfile.docx")
```

You can get and set the metadata properties for a file or folder with `get_item_properties()` and `set_item_properties()`. For the latter, provide the new properties as named arguments to the method. Not all properties can be changed; some, like the file size and last modified date, are read-only. You can also retrieve an object representing the file or folder with `get_item()`, which has methods appropriate for drive items.

```r
od$get_item_properties("Documents/myfile.docx")

# rename a file -- version control via filename is bad, mmkay
od$set_item_properties("Documents/myfile.docx", name="myfile version 2.docx")

# alternatively, you can call the file object's update() method
item <- od$get_item("Documents/myfile.docx")
item$update(name="myfile version 2.docx")
```

To access OneDrive for Business call `get_business_onedrive()`. This also returns an object of class `ms_drive`, so the exact same methods are available as for personal OneDrive.

```r
odb <- get_business_onedrive()

odb$list_items()
odb$open_item("myproject/demo.pptx")
```

## SharePoint

To access a SharePoint site, use the `get_sharepoint_site()` function and provide the site name, URL or ID. You can also list the sites you're following with `list_sharepoint_sites()`.

```r
list_sharepoint_sites()
site <- get_sharepoint_site("My site")
```

The client object has methods to retrieve drives (document libraries) and lists. To show all drives in a site, use the `list_drives()` method, and to retrieve a specific drive, use `get_drive()`. Each drive is an object of class `ms_drive`, just like the OneDrive clients above.

```r
# list of all document libraries under this site
site$list_drives()

# default document library
drv <- site$get_drive()

# same methods as for OneDrive
drv$list_items()
drv$open_item("teamproject/plan.xlsx")
```

To show all lists in a site, use the `get_lists()` method, and to retrieve a specific list, use `get_list()` and supply either the list name or ID.

```r
site$get_lists()

lst <- site$get_list("my-list")
```

You can retrieve the items in a list as a data frame, with `list_items()`. This has arguments `filter` and `select` to do row and column subsetting respectively. `filter` should be an OData expression provided as a string, and `select` should be a string containing a comma-separated list of columns. Any column names in the `filter` expression must be prefixed with `fields/` to distinguish them from item metadata.

```r
# return a data frame containing all list items
lst$list_items()

# get subset of rows and columns
lst$list_items(
    filter="startsWith(fields/firstname, 'John')",
    select="firstname,lastname,title"
)
```

There are also `get_item()`, `create_item()`, `update_item()` and `delete_item()` methods for working directly with individual items.

```r
item <- list$create_item(firstname="Mary", lastname="Smith")
iid <- item$properties$id
list$update_item(iid, firstname="Eliza")
list$delete_item(iid)
```

Finally, you can retrieve subsites with `list_subsites()` and `get_subsite()`. These also return SharePoint site objects, so all the methods above are available for a subsite.

Currently, Microsoft365R only supports SharePoint Online, the cloud-hosted version of the product. Support for SharePoint Server (the on-premises version) may come at a later stage.

## Integration with AzureGraph

In addition to the client functions given above, Microsoft365R enhances the `az_user` and `az_group` classes that are part of AzureGraph, to let you access drives and sites directly from a user or group object.

`az_user` gains `list_drives()` and `get_drive()` methods. The first shows all the drives that the user has access to, including those that are shared from other users. The second retrieves a specific drive, by default the user's OneDrive. Whether these are personal or business drives depends on the tenant that was specified in `AzureGraph::get_graph_login()`/`create_graph_login()`: if the tenant was "consumers", it will be the personal OneDrive.

`az_group` gains `list_drives()`, `get_drive()` and `get_get_sharepoint_site()` methods. The first two do the same as for `az_user`: they retrieve the drive(s) for the group. The third method retrieves the SharePoint site associated with the group, if one exists.

----
<p align="center"><a href="https://github.com/Azure/AzureR"><img src="https://github.com/Azure/AzureR/raw/master/images/logo2.png" width=800 /></a></p>

