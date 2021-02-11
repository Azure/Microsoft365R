app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("OneDrive tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(c("openid", "offline_access"),
    tenant="9188040d-6c67-4c5b-b112-36a304b66dad", app=.microsoft365r_app_id, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("OneDrive tests skipped: unable to login to consumers tenant")

test_that("OneDrive personal works",
{
    od <- get_personal_onedrive()
    expect_is(od, "ms_drive")

    od2 <- get_personal_onedrive(app=app)
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

    expect_silent(od$delete_item(newfolder, confirm=FALSE))
})

