app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
to_addr <- Sys.getenv("AZ_TEST_OUTLOOK_TO_ADDR")
cc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_CC_ADDR")
bcc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_BCC_ADDR")

if(app == "" || to_addr == "" || cc_addr == "" || bcc_addr == "")
    skip("Outlook email attachment tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook email attachment tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(c("https://graph.microsoft.com/Mail.Read", "openid", "offline_access"),
    tenant="9188040d-6c67-4c5b-b112-36a304b66dad", app=.microsoft365r_app_id, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("Outlook email attachment tests skipped: unable to login to consumers tenant")

inbox <- try(call_graph_endpoint(tok, "me/mailFolders/inbox"), silent=TRUE)
if(inherits(inbox, "try-error"))
    skip("Outlook email attachment tests skipped: service not available")

inbox <- try(call_graph_endpoint(tok, "me/mailFolders/inbox"), silent=TRUE)
if(inherits(inbox, "try-error"))
    skip("Outlook email attachment tests skipped: service not available")

fname <- make_name()
folder <- get_personal_outlook()$create_folder(fname)

test_that("Outlook email attachment methods work",
{
    em <- folder$create_email("test email", content_type="html")
    expect_is(em$add_attachment("../resources/logo_small.jpg"), "ms_outlook_email")

    atts <- em$list_attachments()
    expect_is(atts, "list")
    expect_true(!is_empty(atts) && all(sapply(atts, inherits, "ms_outlook_attachment")))

    id1 <- atts[[1]]$properties$id
    att1 <- em$get_attachment("logo_small.jpg")
    expect_is(att1, "ms_outlook_attachment")
    expect_identical(att1$properties$name, "logo_small.jpg")

    dest <- tempfile(fileext=".jpg")
    expect_silent(em$download_attachment("logo_small.jpg", dest=dest, overwrite=TRUE))
    expect_true(files_identical("../resources/logo_small.jpg", dest))

    em$add_attachment("../resources/logo_small.jpg")
    expect_error(em$get_attachment("logo_small.jpg"))  # duplicate filenames

    expect_silent(em$remove_attachment(attachment_id=id1, confirm=FALSE))

    em$add_attachment("https://example.com")
    att2 <- em$get_attachment("example.com")
    expect_is(att2, "ms_outlook_attachment")
    expect_identical(att2$attachment_type, "link")
})


test_that("Attachments from blastula work",
{
    if(!requireNamespace("blastula", quietly=TRUE))
        skip("Blastula tests skipped: package not installed")

    bl_em <- blastula::compose_email(body=blastula::md("## test blastula email"))
    bl_em <- blastula::add_attachment(bl_em, "../resources/logo_small.jpg")

    em <- folder$create_email(bl_em)
    atts <- em$list_attachments()
    expect_is(atts, "list")
    expect_true(!is_empty(atts) && all(sapply(atts, inherits, "ms_outlook_attachment")))
})


test_that("Attachments from emayili work",
{
    if(!requireNamespace("emayili", quietly=TRUE))
        skip("Emayili tests skipped: package not installed")

    ey_em <- emayili::envelope(
        subject="test emayili email",
        html="test emayili email"
    )
    ey_em <- emayili::attachment(ey_em, "../resources/logo_small.jpg")

    em <- folder$create_email(ey_em)
    atts <- em$list_attachments()
    expect_is(atts, "list")
    expect_true(!is_empty(atts) && all(sapply(atts, inherits, "ms_outlook_attachment")))
})


create_bigfile <- function(size)
{
    f <- tempfile()
    con <- file(f, open="wb")
    on.exit(close(con))
    x <- sample(letters, size, replace=TRUE)
    writeBin(charToRaw(paste0(x, collapse="")), con)
    f
}
src <- create_bigfile(4e6)

test_that("Large attachments work",
{
    expect_silent(em <- folder$create_email()$add_attachment(src))
    lst <- em$list_attachments()
    expect_true(!is_empty(lst) && as.numeric(lst[[1]]$properties$size) >= 4e6)
})


test_that("Large attachments from blastula work",
{
    bl_em <- blastula::compose_email("test blastula email")
    bl_em <- blastula::add_attachment(bl_em, src)
    em <- folder$create_email(bl_em)
    lst <- em$list_attachments()
    expect_true(!is_empty(lst) && as.numeric(lst[[1]]$properties$size) >= 4e6)
})


test_that("Large attachments from emayili skipped",
{
    ey_em <- emayili::envelope(text="test emayili email")
    ey_em <- emayili::attachment(ey_em, src)
    expect_warning(em <- folder$create_email(ey_em))
    expect_true(is_empty(em$list_attachments()))
})


test_that("Inline images work",
{
    em <- folder$create_email("test email", content_type="html")
    em$add_image("../resources/logo_small.jpg")
    lst <- em$list_attachments()
    expect_true(!is_empty(lst))
    expect_true(lst[[1]]$properties$isInline)
})


test_that("Inline images from blastula work",
{
    bl_img <- blastula::add_image("../resources/logo_small.jpg")
    bl_em <- blastula::compose_email(blastula::md(c("test blastula email", bl_img)))
    em <- folder$create_email(bl_em)
    lst <- em$list_attachments()
    expect_true(!is_empty(lst))
})


test_that("Links from OneDrive work",
{
    od <- try(get_personal_onedrive(), silent=TRUE)
    if(inherits(od, "try-error"))
        skip("OneDrive attachment tests skipped: service not available")

    f <- write_file()
    item <- od$upload_file(f, basename(f))

    em <- folder$create_email("test email", content_type="html")
    em$add_attachment(item, expiry="14 days", scope="anonymous")
    lst <- em$list_attachments()
    expect_true(!is_empty(lst))

    item$delete(confirm=FALSE)
})


teardown({
    folder$delete(confirm=FALSE)
    unlink(src)
})
