app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(app == "")
    skip("OneDrive tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(c("openid", "offline_access"),
    tenant="9188040d-6c67-4c5b-b112-36a304b66dad", app=.microsoft365r_app_id, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("OneDrive tests skipped: unable to login to consumers tenant")

test_that("OneDrive personal works",
{
    od <- get_personal_onedrive()
    expect_is(od, "ms_drive")

    od2 <- get_personal_onedrive(app=app)
    expect_is(od2, "ms_drive")

    ls <- od$list_items()
    expect_is(ls, "data.frame")

    newfolder <- make_name()
    expect_silent(od$create_folder(newfolder))

    src <- "../resources/file.json"
    dest <- file.path(newfolder, basename(src))
    newsrc <- tempfile()
    expect_silent(od$upload_file(src, dest))
    expect_silent(od$download_file(dest, newsrc))

    expect_true(files_identical(src, newsrc))

    item <- od$get_item(dest)
    expect_is(item, "ms_drive_item")

    expect_silent(od$set_item_properties(dest, name="newname"))
    expect_silent(item$sync_fields())
    expect_identical(item$properties$name, "newname")
    expect_silent(item$update(name=basename(dest)))
    expect_identical(item$properties$name, basename(dest))

    expect_silent(od$delete_item(newfolder, confirm=FALSE))
})


test_that("Drive item methods work",
{
    od <- get_personal_onedrive()

    root <- od$get_item("/")
    expect_is(root, "ms_drive_item")

    tmpname1 <- make_name(10)
    folder1 <- root$create_folder(tmpname1)
    expect_is(folder1, "ms_drive_item")
    expect_true(folder1$is_folder())

    tmpname2 <- make_name(10)
    folder2 <- folder1$create_folder(tmpname2)
    expect_is(folder2, "ms_drive_item")
    expect_true(folder2$is_folder())

    src <- write_file()
    expect_silent(file1 <- root$upload(src))
    expect_is(file1, "ms_drive_item")
    expect_false(file1$is_folder())
    expect_error(file1$create_folder("bad"))

    file1_0 <- root$get_item(basename(src))
    expect_is(file1_0, "ms_drive_item")
    expect_false(file1_0$is_folder())
    expect_identical(file1_0$properties$name, file1$properties$name)

    dest1 <- tempfile()
    expect_silent(file1$download(dest1))
    expect_true(files_identical(src, dest1))

    expect_silent(file2 <- folder1$upload(src))
    expect_is(file2, "ms_drive_item")

    dest2 <- tempfile()
    expect_silent(file2$download(dest2))
    expect_true(files_identical(src, dest2))

    dest3 <- tempfile()
    expect_silent(file3 <- folder2$upload(src, basename(dest3)))
    expect_is(file3, "ms_drive_item")
    expect_identical(file3$properties$name, basename(dest3))
    expect_silent(file3$download(dest3))
    expect_true(files_identical(src, dest3))

    file3_1 <- folder1$get_item(file.path(tmpname2, basename(dest3)))
    expect_is(file3_1, "ms_drive_item")
    expect_identical(file3_1$properties$name, file3$properties$name)

    lst0 <- root$list_files()
    expect_is(lst0, "data.frame")
    lst0_f <- root$list_files(info="name", full_names=TRUE)
    expect_type(lst0_f, "character")
    expect_true(all(substr(lst0_f, 1, 1) == "/"))

    lst0_1 <- root$list_files(tmpname1)
    lst1 <- folder1$list_files()
    expect_identical(lst0_1, lst1)

    lst1_f <- folder1$list_files(tmpname2, info="name", full_names=TRUE)
    expect_type(lst1_f, "character")
    expect_true(all(grepl(paste0("^", tmpname2), lst1_f)))

    expect_silent(file1$delete(confirm=FALSE))
    expect_silent(folder1$delete(confirm=FALSE))
})


test_that("Methods work with filenames with special characters",
{
    od <- get_personal_onedrive()

    test_name <- paste(make_name(5), "plus spaces and áccénts")
    src <- write_file(fname=file.path(tempdir(), test_name))

    expect_silent(od$upload_file(src, basename(src)))
    expect_silent(item <- od$get_item(basename(test_name)))
    expect_true(item$properties$name == basename(test_name))
    expect_silent(item$delete(confirm=FALSE))
})

