tenant <- "consumers"
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
from_addr <- Sys.getenv("AZ_TEST_OUTLOOK_FROM_ADDR")
to_addr <- Sys.getenv("AZ_TEST_OUTLOOK_TO_ADDR")
cc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_CC_ADDR")
bcc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_BCC_ADDR")

if(app == "" || from_addr == "" || to_addr == "" || cc_addr == "" || bcc_addr == "")
    skip("Outlook email tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook email tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Mail.Send", "Mail.ReadWrite", "User.Read"))
if(is.null(tok))
    skip("Outlook tests skipped: unable to login to consumers tenant")

inbox <- try(call_graph_endpoint(tok, "me/mailFolders/inbox"), silent=TRUE)
if(inherits(inbox, "try-error"))
    skip("Outlook tests skipped: service not available")

get_to_addr <- function(x, n=1) x$properties$toRecipients[[n]]$emailAddress$address
get_cc_addr <- function(x, n=1) x$properties$ccRecipients[[n]]$emailAddress$address
get_bcc_addr <- function(x, n=1) x$properties$bccRecipients[[n]]$emailAddress$address
get_replyto_addr <- function(x, n=1) x$properties$replyTo[[n]]$emailAddress$address

fname <- make_name()
folder <- get_personal_outlook(token=tok)$create_folder(fname)

test_that("Outlook email methods work",
{
    em <- folder$create_email()
    expect_is(em, c("ms_outlook_email", "ms_outlook_object"))

    expect_identical(em$properties$body$content, "")
    expect_identical(em$properties$subject, "")
    expect_true(is_empty(em$properties$toRecipients))
    expect_true(is_empty(em$properties$ccRecipients))
    expect_true(is_empty(em$properties$bccRecipients))

    body_text <- "test message body"
    body_html <- "<p>test html message body</p>"
    subj <- "test subject line"

    em$set_body(body_text)
    expect_identical(em$properties$body$content, body_text)

    em$set_body(body_html, "html")
    expect_true(grepl(body_html, em$properties$body$content))
    expect_identical(em$properties$body$contentType, "html")

    em$set_subject(subj)
    expect_identical(em$properties$subject, subj)

    em$add_recipients(to_addr)
    expect_identical(get_to_addr(em), to_addr)

    em$add_recipients(cc=cc_addr)
    expect_identical(get_to_addr(em), to_addr)
    expect_identical(get_cc_addr(em), cc_addr)

    em$add_recipients(cc=bcc_addr)
    expect_identical(get_cc_addr(em), cc_addr)
    expect_identical(get_cc_addr(em, 2), bcc_addr)

    em$set_recipients(to_addr, cc=cc_addr, bcc=bcc_addr)
    expect_identical(get_to_addr(em), to_addr)
    expect_identical(get_cc_addr(em), cc_addr)
    expect_error(get_cc_addr(em, 2))
    expect_identical(get_bcc_addr(em), bcc_addr)

    em$set_reply_to(from_addr)
    expect_identical(get_replyto_addr(em), from_addr)
    expect_identical(get_to_addr(em), to_addr)
    expect_identical(get_cc_addr(em), cc_addr)
    expect_identical(get_bcc_addr(em), bcc_addr)
})


test_that("Creating email with blastula works",
{
    if(!requireNamespace("blastula", quietly=TRUE))
        skip("Blastula tests skipped: package not installed")

    bl_em <- blastula::compose_email(body=blastula::md("## test blastula email"))
    em <- folder$create_email(bl_em)

    expect_identical(em$properties$body$contentType, "html")
    expect_true(grepl("test blastula email", em$properties$body$content))
})


test_that("Creating email with emayili works",
{
    if(!requireNamespace("emayili", quietly=TRUE))
        skip("Emayili tests skipped: package not installed")

    ey_em <- emayili::envelope(
        to=to_addr,
        subject="test emayili email",
        html="<p>test emayili email</p>"
    )
    em <- folder$create_email(ey_em)

    expect_identical(em$properties$body$contentType, "html")
    expect_true(grepl("test emayili email", em$properties$body$content))
    expect_identical(em$properties$subject, "test emayili email")
    expect_identical(get_to_addr(em), to_addr)
})


teardown({
    folder$delete(confirm=FALSE)
})
