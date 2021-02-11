tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(tenant == "" || app == "")
    skip("OneDrive for Business tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive for Business tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(
    c("https://graph.microsoft.com/.default",
      "openid",
      "offline_access"),
    tenant=tenant, app=app, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("OneDrive for Business tests skipped: no access to tenant")

test_that("OneDrive for Business works",
{
    gr <- AzureGraph::ms_graph$new(token=tok)
    drv <- try(gr$get_user()$get_drive(), silent=TRUE)
    if(inherits(drv, "try-error"))
        skip("OneDrive for Business tests skipped: service not available")

    od <- get_business_onedrive(tenant=tenant)
    expect_is(od, "ms_drive")

    od2 <- get_business_onedrive(tenant=tenant, app=app)
    expect_is(od2, "ms_drive")

    ls <- od$list_items()
    expect_is(ls, "data.frame")

    newfolder <- make_name()
    expect_silent(od$create_folder(newfolder))

    src <- "../resources/file.json"
    dest <- file.path(newfolder, basename(src))
    newsrc <- tempfile()
    expect_silent(od$upload_file(src, dest))
    expect_silent(od$download_file(dest, newsrc))

    expect_true(files_identical(src, newsrc))

    item <- od$get_item(dest)
    expect_is(item, "ms_drive_item")

    expect_silent(od$set_item_properties(dest, name="newname"))
    expect_silent(item$sync_fields())
    expect_identical(item$properties$name, "newname")
    expect_silent(item$update(name=basename(dest)))
    expect_identical(item$properties$name, basename(dest))

    # ODB requires that folders be empty before deleting them (!)
    expect_silent(item$delete(confirm=FALSE))

    expect_silent(od$delete_item(newfolder, confirm=FALSE))
})

