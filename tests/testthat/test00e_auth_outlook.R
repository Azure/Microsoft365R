tenant <- "consumers"
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("Outlook authentication tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook authentication tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Mail.Send", "Mail.ReadWrite", "User.Read"))
if(is.null(tok))
    skip("Outlook authentication tests skipped: unable to login to consumers tenant")

inbox <- try(call_graph_endpoint(tok, "me/mailFolders/inbox"), silent=TRUE)
if(inherits(inbox, "try-error"))
    skip("Outlook authentication tests skipped: service not available")

test_that("Outlook authentication works",
{
    outl <- get_personal_outlook(token=tok)
    expect_is(outl, "ms_outlook")

    outl2 <- get_personal_outlook(app=app)
    expect_is(outl2, "ms_outlook")
    expect_identical(outl$properties$id, outl2$properties$id)
})
