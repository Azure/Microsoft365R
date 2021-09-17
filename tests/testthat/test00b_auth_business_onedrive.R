tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(tenant == "" || app == "")
    skip("OneDrive for Business authentication tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive for Business authentication tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Files.ReadWrite.All", "User.Read"))
if(is.null(tok))
    skip("OneDrive for Business authentication tests skipped: no access to tenant")

drv <- try(call_graph_endpoint(tok, "me/drive"), silent=TRUE)
if(inherits(drv, "try-error"))
    skip("OneDrive for Business tests skipped: service not available")

test_that("OneDrive authentication works",
{
    drv <- get_business_onedrive(token=tok)
    expect_is(drv, "ms_drive")

    drv2 <- get_business_onedrive(tenant=tenant, app=app)
    expect_is(drv2, "ms_drive")
    expect_identical(drv$properties$id, drv2$properties$id)
})
