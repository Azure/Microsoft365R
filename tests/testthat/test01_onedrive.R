tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(tenant == "" || app == "")
    skip("Authentication tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Authentication tests skipped: must be in interactive session")


test_that("OneDrive personal works",
{
    od <- personal_onedrive()
    expect_is(od, "ms_drive")

    ls <- od$list_items()
    expect_is(ls, "data.frame")

    newfolder <- make_name()
    expect_silent(od$create_folder(newfolder))

    src <- write_file()
    dest <- file.path(newfolder, basename(src))
    newsrc <- tempfile()
    expect_silent(od$upload_file(src, dest))
    expect_silent(od$download_file(dest, newsrc))

    expect_true(files_identical(src, newsrc))

    expect_silent(od$delete_item(newfolder, confirm=FALSE))
})

