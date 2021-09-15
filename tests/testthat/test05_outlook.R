tenant <- "consumers"
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("Outlook tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Mail.Send", "Mail.ReadWrite", "User.Read"))
if(is.null(tok))
    skip("Outlook tests skipped: unable to login to consumers tenant")

inbox <- try(call_graph_endpoint(tok, "me/mailFolders/inbox"), silent=TRUE)
if(inherits(inbox, "try-error"))
    skip("Outlook tests skipped: service not available")

outl <- get_personal_outlook(token=tok)

test_that("Outlook client works",
{
    expect_is(outl, c("ms_outlook", "ms_outlook_object"))

    folders <- outl$list_folders()
    expect_is(folders, "list")
    expect_true(all(sapply(folders, inherits, "ms_outlook_folder")))

    f1 <- folders[[1]]$properties$displayName
    fpager <- outl$list_folders(filter=sprintf("displayName eq '%s'", f1), n=NULL)
    expect_is(fpager, "ms_graph_pager")
    folders1 <- fpager$value
    expect_true(length(folders1) ==1 && inherits(folders1[[1]], "ms_outlook_folder"))

    emails <- outl$list_emails()
    expect_is(emails, "list")
    expect_true(all(sapply(emails, inherits, "ms_outlook_email")))

    subj1 <- emails[[1]]$properties$subject
    empager <- outl$list_emails(filter=sprintf("subject eq '%s'", subj1), n=NULL)
    expect_is(empager, "ms_graph_pager")
    emails1 <- empager$value
    expect_true(length(emails1) == 1 && inherits(emails1[[1]], "ms_outlook_email"))

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



