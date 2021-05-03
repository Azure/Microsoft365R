tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
site_name <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_NAME")
site_url <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_URL")
site_id <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_ID")
list_name <- Sys.getenv("AZ_TEST_SHAREPOINT_LIST_NAME")
list_id <- Sys.getenv("AZ_TEST_SHAREPOINT_LIST_ID")

if(tenant == "" || app == "" || site_name == "" || site_url == "" || site_id == "" || list_name == "" || list_id == "")
    skip("SharePoint tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("SharePoint tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(
    c("https://graph.microsoft.com/.default",
      "openid",
      "offline_access"),
    tenant=tenant, app=app, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("SharePoint tests skipped: no access to tenant")

site <- try(call_graph_endpoint(tok, file.path("sites", site_id)), silent=TRUE)
if(inherits(site, "try-error"))
    skip("SharePoint tests skipped: service not available")

test_that("SharePoint client works",
{
    expect_error(get_sharepoint_site(site_name=site_name, site_url=site_url, site_id=site_id,
                                     tenant=tenant, app=app))

    site1 <- get_sharepoint_site(site_name=site_name, tenant=tenant, app=app)
    expect_is(site1, "ms_site")
    expect_identical(site1$properties$displayName, site_name)

    site2 <- get_sharepoint_site(site_url=site_url, tenant=tenant, app=app)
    expect_is(site2, "ms_site")
    expect_identical(site1$properties$webUrl, site_url)

    site3 <- get_sharepoint_site(site_id=site_id, tenant=tenant, app=app)
    expect_is(site3, "ms_site")
    expect_identical(site1$properties$id, site_id)

    expect_identical(site1$properties, site2$properties)
    expect_identical(site2$properties, site3$properties)

    sites <- list_sharepoint_sites()
    expect_is(sites, "list")
    expect_true(all(sapply(sites, inherits, "ms_site")))
})

test_that("SharePoint methods work",
{
    site <- get_sharepoint_site(site_name, tenant=tenant, app=app)
    expect_is(site, "ms_site")

    # drive functionality tested in test02
    drives <- site$list_drives()
    expect_is(drives, "list")
    expect_true(all(sapply(drives, inherits, "ms_drive")))

    # filtering not yet supported for drives
    # drvpager <- site$list_drives(filter="name eq 'Documents'", n=NULL)
    # expect_is(drvpager, "ms_graph_pager")
    # drv0 <- drvpager$value
    # expect_is(drv0, "list")
    # expect_true(length(drv0) == 1 && inherits(drv0[[1]], "ms_drive"))

    drv <- site$get_drive()
    expect_is(drv, "ms_drive")

    grp <- site$get_group()
    expect_is(grp, "az_group")

    # list
    lists <- site$get_lists()
    expect_is(lists, "list")
    expect_true(all(sapply(lists, inherits, "ms_list")))

    # filtering not yet supported
    # lstpager <- site$get_lists(filter=sprintf("displayName eq '%s'", list_name), n=NULL)
    # expect_is(lstpager, "ms_graph_pager")
    # lst0 <- lstpager$value
    # expect_true(length(lst0) == 1 && inherits(lst0[[1]], "ms_list"))

    lst <- site$get_list(list_name=list_name)
    lst2 <- site$get_list(list_id=list_id)
    expect_is(lst, "ms_list")
    expect_is(lst2, "ms_list")
    expect_identical(lst$properties, lst2$properties)

    cols <- lst$get_column_info()
    expect_is(cols, "data.frame")

    items <- lst$list_items()
    expect_is(items, "data.frame")

    items2 <- lst$list_items(all_metadata=TRUE)
    expect_is(items2, "data.frame")
    expect_identical(items, items2$fields)

    items3 <- lst$list_items(as_data_frame=FALSE)
    expect_is(items3, "list")
    expect_true(all(sapply(items3, inherits, "ms_list_item")))

    itpager <- lst$list_items(filter=sprintf("fields/Title eq '%s'", items3[[1]]$properties$fields$Title), n=NULL)
    expect_is(itpager, "ms_graph_pager")
    items0 <- itpager$value
    expect_true(is.data.frame(items0) && nrow(items0) == 1)

    item_id <- items3[[1]]$properties$id
    item <- lst$get_item(item_id)
    expect_is(item, "ms_list_item")
    expect_false(is_empty(item$properties$fields))

    newtitle <- make_name(10)
    newitem <- lst$create_item(Title=newtitle)
    expect_is(newitem, "ms_list_item")
    newid <- newitem$properties$id

    updatetitle <- make_name(10)
    expect_silent(lst$update_item(newid, Title=updatetitle))

    updateitem <- lst$get_item(newid)
    expect_identical(updateitem$properties$fields$Title, updatetitle)

    expect_silent(lst$delete_item(newid, confirm=FALSE))
    items4 <- lst$list_items()
    expect_identical(nrow(items), nrow(items4))

    df <- data.frame(Title=c("item1", "item2", "item3"), stringsAsFactors=FALSE)
    items5 <- lst$bulk_import(df)
    expect_is(items5, "list")
    expect_true(all(sapply(items5, inherits, "ms_list_item")))

    expect_silent(lapply(items5, function(it) it$delete(confirm=FALSE)))
})
