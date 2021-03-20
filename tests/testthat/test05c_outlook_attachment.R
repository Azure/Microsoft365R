app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
to_addr <- Sys.getenv("AZ_TEST_OUTLOOK_TO_ADDR")
cc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_CC_ADDR")
bcc_addr <- Sys.getenv("AZ_TEST_OUTLOOK_BCC_ADDR")

if(app == "" || to_addr == "" || cc_addr == "" || bcc_addr == "")
    skip("Outlook email attachment tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Outlook email attachment tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(c("openid", "offline_access"),
    tenant="9188040d-6c67-4c5b-b112-36a304b66dad", app=.microsoft365r_app_id, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("Outlook tests skipped: unable to login to consumers tenant")

fname <- make_name()
folder <- get_personal_outlook()$create_folder(fname)

test_that("Outlook email attachment methods work",
{
    em <- folder$create_email()
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


teardown({
    folder$delete(confirm=FALSE)
})