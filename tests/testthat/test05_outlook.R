app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("Outlook tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(c("https://graph.microsoft.com/Mail.Read", "openid", "offline_access"),
    tenant="9188040d-6c67-4c5b-b112-36a304b66dad", app=.microsoft365r_app_id, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("Outlook tests skipped: unable to login to consumers tenant")

inbox <- try(call_graph_endpoint(tok, "me/mailFolders/inbox"), silent=TRUE)
if(inherits(inbox, "try-error"))
    skip("Outlook tests skipped: service not available")

test_that("Outlook client works",
{
    outl <- get_personal_outlook()
    expect_is(outl, c("ms_outlook", "ms_outlook_object"))

    outl2 <- get_personal_outlook(app=app)
    expect_is(outl2, c("ms_outlook", "ms_outlook_object"))

    folders <- outl$list_folders()
    expect_is(folders, "list")
    expect_true(all(sapply(folders, inherits, "ms_outlook_folder")))

    emails <- outl$list_emails()
    expect_is(emails, "list")
    expect_true(all(sapply(emails, inherits, "ms_outlook_email")))

    f1name <- make_name()
    f1 <- outl$create_folder(f1name)
    expect_is(f1, "ms_outlook_folder")

    expect_error(outl$create_folder(f1name))
    f11 <- outl$get_folder(f1name)
    expect_identical(f1$properties$id, f11$properties$id)
    expect_silent(outl$delete_folder(f1name, confirm=FALSE))

    fnames <- sapply(outl$list_folders(), function(x) x$properties$displayName)
    expect_false(f1name %in% fnames)

    expect_is(outl$get_inbox(), "ms_outlook_folder")
    expect_is(outl$get_drafts(), "ms_outlook_folder")
    expect_is(outl$get_sent_items(), "ms_outlook_folder")
    expect_is(outl$get_deleted_items(), "ms_outlook_folder")

    eml <- outl$create_email("hello from R")
    expect_is(eml, "ms_outlook_email")
    expect_silent(eml$delete(confirm=FALSE))
})



