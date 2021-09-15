app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("Authentication tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive tests skipped: must be in interactive session")


scopes <- c("Files.ReadWrite.All", "User.Read")

test_that("Authenticating works with consumers tenant",
{
    tenant <- "consumers"

    tok <- get_test_token(tenant, app, scopes)
    if(is.null(tok))
        skip("Consumer tenant tests skipped: unable to login to consumers tenant")

    drv <- try(call_graph_endpoint(tok, "me/drive"), silent=TRUE)
    if(inherits(drv, "try-error"))
        skip("Consumer tenant tests skipped: service not available")

    expect_type(drv, "list")
    expect_true("id" %in% names(drv) && is.character(drv$id) && nchar(drv$id) > 0)

    gr <- do_login(tenant, app, scopes, NULL)
    expect_is(gr, "ms_graph")

    drv2 <- gr$call_graph_endpoint("me/drive")

    expect_type(drv2, "list")
    expect_true("id" %in% names(drv2) && is.character(drv2$id) && nchar(drv2$id) > 0)
    expect_identical(drv$properties$id, drv2$properties$id)
})


test_that("Authenticating works with org tenant",
{
    tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
    if(tenant == "")
        skip("Authentication tests skipped: Microsoft Graph credentials not set")

    tok <- get_test_token(tenant, app, scopes)
    if(is.null(tok))
        skip("Org tenant tests skipped: unable to login to consumers tenant")

    drv <- try(call_graph_endpoint(tok, "me/drive"), silent=TRUE)
    if(inherits(drv, "try-error"))
        skip("Org tenant tests skipped: service not available")

    expect_type(drv, "list")
    expect_true("id" %in% names(drv) && is.character(drv$id) && nchar(drv$id) > 0)

    gr <- do_login(tenant, app, scopes, NULL)
    expect_is(gr, "ms_graph")

    drv2 <- gr$call_graph_endpoint("me/drive")

    expect_type(drv2, "list")
    expect_true("id" %in% names(drv2) && is.character(drv2$id) && nchar(drv2$id) > 0)
    expect_identical(drv$properties$id, drv2$properties$id)
})

