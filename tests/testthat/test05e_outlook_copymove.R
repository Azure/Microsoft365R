tenant <- "consumers"
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
from_addr <- Sys.getenv("AZ_TEST_OUTLOOK_FROM_ADDR")
to_addr <- Sys.getenv("AZ_TEST_OUTLOOK_TO_ADDR")
cc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_CC_ADDR")
bcc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_BCC_ADDR")

if(app == "" || from_addr == "" || to_addr == "" || cc_addr == "" || bcc_addr == "")
    skip("Outlook email tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook email send tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Mail.Send", "Mail.ReadWrite", "User.Read"))
if(is.null(tok))
    skip("Outlook tests send skipped: unable to login to consumers tenant")

get_to_addr <- function(x, n=1) x$properties$toRecipients[[n]]$emailAddress$address
get_cc_addr <- function(x, n=1) x$properties$ccRecipients[[n]]$emailAddress$address
get_bcc_addr <- function(x, n=1) x$properties$bccRecipients[[n]]$emailAddress$address
get_replyto_addr <- function(x, n=1) x$properties$replyTo[[n]]$emailAddress$address

outl <- get_personal_outlook(token=tok)
srcname <- make_name()
destname <- make_name()
src <- outl$create_folder(srcname)
dest <- outl$create_folder(destname)

test_that("Outlook email copy/move methods work",
{
    em <- src$create_email()
    expect_is(em, "ms_outlook_email")

    expect_error(em$copy(destname))  # must supply an object, not a name or ID

    expect_silent(em2 <- em$copy(dest))
    expect_identical(em2$properties$parentFolderId, dest$properties$id)
    expect_identical(em$properties$parentFolderId, src$properties$id)
    em$sync_fields()  # this should be a no-op
    expect_identical(em$properties$parentFolderId, src$properties$id)

    expect_silent(em3 <- em$move(dest))
    expect_identical(em3$properties$parentFolderId, dest$properties$id)
    expect_identical(em$properties$parentFolderId, dest$properties$id)
})


test_that("Outlook folder copy/move methods work",
{
    fol <- src$create_folder(make_name())
    expect_is(fol, "ms_outlook_folder")

    expect_error(fol$copy(destname))  # must supply an object, not a name or ID

    expect_silent(fol2 <- fol$copy(dest))
    expect_identical(fol2$properties$parentFolderId, dest$properties$id)
    expect_identical(fol$properties$parentFolderId, src$properties$id)
    fol$sync_fields()  # this should be a no-op
    expect_identical(fol$properties$parentFolderId, src$properties$id)

    fol3 <- src$create_folder(make_name())
    expect_silent(fol4 <- fol3$move(dest))
    expect_identical(fol4$properties$parentFolderId, dest$properties$id)
    expect_identical(fol3$properties$parentFolderId, dest$properties$id)
})


teardown({
    src$delete(confirm=FALSE)
    dest$delete(confirm=FALSE)
})
