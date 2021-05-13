app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
from_addr <- Sys.getenv("AZ_TEST_OUTLOOK_FROM_ADDR")
to_addr <- Sys.getenv("AZ_TEST_OUTLOOK_TO_ADDR")
cc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_CC_ADDR")
bcc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_BCC_ADDR")

if(app == "" || from_addr == "" || to_addr == "" || cc_addr == "" || bcc_addr == "")
    skip("Outlook email tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook email send tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(c("https://graph.microsoft.com/Mail.Read", "openid", "offline_access"),
    tenant="9188040d-6c67-4c5b-b112-36a304b66dad", app=.microsoft365r_app_id, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("Outlook tests send skipped: unable to login to consumers tenant")

inbox <- try(call_graph_endpoint(tok, "me/mailFolders/inbox"), silent=TRUE)
if(inherits(inbox, "try-error"))
    skip("Outlook tests skipped: service not available")

get_to_addr <- function(x, n=1) x$properties$toRecipients[[n]]$emailAddress$address
get_cc_addr <- function(x, n=1) x$properties$ccRecipients[[n]]$emailAddress$address
get_bcc_addr <- function(x, n=1) x$properties$bccRecipients[[n]]$emailAddress$address
get_replyto_addr <- function(x, n=1) x$properties$replyTo[[n]]$emailAddress$address

fname <- make_name()
outl <- get_personal_outlook()
folder <- outl$create_folder(fname)

test_that("Outlook email sending methods work",
{
    subj <- paste("test send", make_name(10))
    em <- outl$create_email("test send",
        subject=subj,
        to=to_addr,
        cc=cc_addr,
        reply_to=from_addr)
    expect_silent(em$send())

    Sys.sleep(5)
    expect_silent(lst <- outl$list_emails())

    wch <- which(sapply(lst, function(em) em$properties$subject == subj))
    expect_true(!is_empty(wch))

    em2 <- lst[[wch[1]]]
    expect_identical(get_replyto_addr(em2), from_addr)

    expect_silent(em2$create_reply("test reply")$send())

    # reply-all ignores our own addresses! must add at least one address manually for sending to work
    ra <- em2$create_reply_all("test reply all")
    ra$add_recipients(to=to_addr, cc=cc_addr)
    expect_silent(ra$send())

    fw <- em2$create_forward("test forward", to=to_addr, cc=cc_addr)
    expect_silent(fw$send())
})


teardown({
    folder$delete(confirm=FALSE)
})
