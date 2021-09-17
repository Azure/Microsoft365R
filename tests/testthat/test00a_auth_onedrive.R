tenant <- "consumers"
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("OneDrive authentication tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive authentication tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Files.ReadWrite.All", "User.Read"))
if(is.null(tok))
    skip("OneDrive authentication tests skipped: unable to login to consumers tenant")

test_that("OneDrive authentication works",
{
    drv <- get_personal_onedrive(token=tok)
    expect_is(drv, "ms_drive")

    drv2 <- get_personal_onedrive(app=app)
    expect_is(drv2, "ms_drive")
    expect_identical(drv$properties$id, drv2$properties$id)
})
