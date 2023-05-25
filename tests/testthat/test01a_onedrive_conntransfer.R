tenant <- "consumers"
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("OneDrive tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Files.ReadWrite.All", "User.Read"))
if(is.null(tok))
    skip("OneDrive tests skipped: unable to login to consumers tenant")

drv <- try(call_graph_endpoint(tok, "me/drive"), silent=TRUE)
if(inherits(drv, "try-error"))
    skip("OneDrive tests skipped: service not available")

opt_use_itemid <- options(microsoft365r_use_itemid_in_path=TRUE)
od <- ms_drive$new(tok, tenant, drv)
folder <- od$create_folder(make_name())

test_that("OneDrive file transfer extras work",
{
    expect_is(od, "ms_drive")

    src <- "../resources/file.json"
    img <- "../resources/logo_small.jpg"

    # upload raw connection
    r <- readBin(img, what="raw", n=file.size(img))
    rcon <- rawConnection(r)
    expect_silent(folder$upload(rcon, "raw.jpg"))

    # upload text connection
    tcon <- textConnection(readLines(src))
    expect_silent(folder$upload(tcon, "text.json"))

    # download raw vector
    expect_silent(rret <- folder$get_item("raw.jpg")$download(NULL))
    expect_type(rret, "raw")
    expect_identical(r, rret)

    expect_silent(tret <- folder$get_item("text.json")$download(NULL))
    expect_type(tret, "raw")
    expect_identical(rawToChar(tret), paste0(readLines(src), collapse="\n"))
})


teardown({
    options(opt_use_itemid)
    folder$delete(confirm=FALSE)
})
