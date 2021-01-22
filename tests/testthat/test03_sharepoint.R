tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
site_url <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_URL")
site_id <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_ID")
list_name <- Sys.getenv("AZ_TEST_SHAREPOINT_LIST_NAME")
list_id <- Sys.getenv("AZ_TEST_SHAREPOINT_LIST_ID")

if(tenant == "" || app == "" || site_url == "" || site_id == "" || list_name == "" || list_id == "")
    skip("SharePoint tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive for Business tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(
    c("https://graph.microsoft.com/.default",
      "openid",
      "offline_access"),
    tenant=tenant, app=app, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("SharePoint tests skipped: no access to tenant")

test_that("SharePoint client works",
{
    gr <- AzureGraph::ms_graph$new(token=tok)
    testsite <- try(gr$call_graph_endpoint(file.path("sites", site_id)), silent=TRUE)
    if(inherits(testsite, "try-error"))
        skip("SharePoint tests skipped: service not available")

    site <- sharepoint_site(site_url, tenant=tenant, app=app)
    expect_is(site, "ms_site")

    site2 <- sharepoint_site(site_id=site_id, tenant=tenant, app=app)
    expect_is(site2, "ms_site")
    expect_identical(site$properties, site2$properties)

    # drive -- functionality tested in test02
    drives <- site$list_drives()
    expect_is(drives, "list")
    expect_true(all(sapply(drives, inherits, "ms_drive")))

    drv <- site$get_drive()
    expect_is(drv, "ms_drive")

    # list
    lists <- site$get_lists()
    expect_is(lists, "list")
    expect_true(all(sapply(lists, inherits, "ms_sharepoint_list")))

    lst <- site$get_list(list_name=list_name)
    lst2 <- site$get_list(list_id=list_id)
    expect_is(lst, "ms_sharepoint_list")
    expect_is(lst2, "ms_sharepoint_list")
    expect_identical(lst$properties, lst2$properties)

    cols <- lst$get_column_info()
    expect_is(cols, "data.frame")

    items <- lst$list_items()
    expect_is(items, "data.frame")

    items2 <- lst$list_items(all_metadata=TRUE)
    expect_is(items2, "data.frame")
    expect_identical(items, items2$fields)
})
