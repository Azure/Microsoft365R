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

test_that("OneDrive item copy/move methods work",
{
    expect_is(folder, "ms_drive_item")

    newfoldername <- make_name()
    expect_is(newfolder <- od$create_folder(newfoldername), "ms_drive_item")

    src <- "../resources/file.json"
    expect_silent(folder$upload(src))

    odsrc <- file.path(folder$properties$name, basename(src))
    # copy via path
    destpath1 <- file.path(newfoldername, "copy1.json")
    expect_silent(od$copy_item(odsrc, dest=destpath1))

    # wait for async copy
    Sys.sleep(2)
    expect_is(od$get_item(destpath1), "ms_drive_item")

    # copy via name
    expect_silent(od$copy_item(odsrc, dest="copy2.json", dest_folder_item=newfolder))

    # wait for async copy
    destpath2 <- file.path(newfoldername, "copy2.json")
    Sys.sleep(2)
    expect_is(od$get_item(destpath2), "ms_drive_item")

    # move via path
    it1 <- od$move_item(odsrc, dest=file.path(newfoldername, "move1.json"))
    expect_identical(it1$properties$name, "move1.json")
    expect_error(folder$get_item(basename(src)))

    # move via name
    newsrc <- file.path(newfoldername, "move1.json")
    it2 <- od$move_item(newsrc, dest="move2.json", dest_folder_item=newfolder)
    expect_identical(it2$properties$name, "move2.json")
    expect_error(newfolder$get_item("move1.json"))
})


teardown({
    options(opt_use_itemid)
})
