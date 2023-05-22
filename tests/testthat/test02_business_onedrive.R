tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")

if(tenant == "" || app == "")
    skip("OneDrive for Business tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("OneDrive for Business tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Files.ReadWrite.All", "User.Read"))
if(is.null(tok))
    skip("OneDrive for Business tests skipped: no access to tenant")

drv <- try(call_graph_endpoint(tok, "me/drive"), silent=TRUE)
if(inherits(drv, "try-error"))
    skip("OneDrive for Business tests skipped: service not available")

od <- ms_drive$new(tok, tenant, drv)

test_that("OneDrive for Business works",
{
    expect_is(od, "ms_drive")

    lst <- od$list_items()
    expect_is(lst, "data.frame")

    newfolder <- make_name()
    expect_silent(od$create_folder(newfolder))

    src <- "../resources/file.json"
    dest <- file.path(newfolder, basename(src))
    newsrc <- tempfile()
    expect_silent(od$upload_file(src, dest))
    expect_silent(od$download_file(dest, dest=newsrc))

    expect_true(files_identical(src, newsrc))

    item <- od$get_item(dest)
    expect_is(item, "ms_drive_item")

    pager <- od$list_files(newfolder, filter=sprintf("name eq '%s'", basename(src)), n=NULL)
    expect_is(pager, "ms_graph_pager")
    lst1 <- pager$value
    expect_is(lst1, "data.frame")
    expect_identical(nrow(lst1), 1L)

    expect_silent(od$set_item_properties(dest, name="newname"))
    expect_silent(item$sync_fields())
    expect_identical(item$properties$name, "newname")
    expect_silent(item$update(name=basename(dest)))
    expect_identical(item$properties$name, basename(dest))

    # ODB requires that folders be empty before deleting them (!)
    expect_silent(item$delete(confirm=FALSE))

    expect_silent(od$delete_item(newfolder, confirm=FALSE))
})


test_that("Drive item methods work",
{
    root <- od$get_item("/")
    expect_is(root, "ms_drive_item")

    rootp <- root$get_parent_folder()
    expect_is(rootp, "ms_drive_item")
    expect_equal(rootp$properties$name, "root")

    tmpname1 <- make_name(10)
    folder1 <- root$create_folder(tmpname1)
    expect_is(folder1, "ms_drive_item")
    expect_true(folder1$is_folder())

    folder1p <- folder1$get_parent_folder()
    expect_equal(rootp$properties$name, "root")

    tmpname2 <- make_name(10)
    folder2 <- folder1$create_folder(tmpname2)
    expect_is(folder2, "ms_drive_item")
    expect_true(folder2$is_folder())

    folder2p <- folder2$get_parent_folder()
    expect_equal(folder2p$properties$name, folder1$properties$name)

    src <- write_file()
    expect_silent(file1 <- root$upload(src))
    expect_is(file1, "ms_drive_item")
    expect_false(file1$is_folder())
    expect_error(file1$create_folder("bad"))

    file1p <- file1$get_parent_folder()
    expect_equal(file1p$properties$name, "root")

    file1_0 <- root$get_item(basename(src))
    expect_is(file1_0, "ms_drive_item")
    expect_false(file1_0$is_folder())
    expect_identical(file1_0$properties$name, file1$properties$name)

    dest1 <- tempfile()
    expect_silent(file1$download(dest1))
    expect_true(files_identical(src, dest1))

    expect_silent(file2 <- folder1$upload(src))
    expect_is(file2, "ms_drive_item")

    file2p <- file2$get_parent_folder()
    expect_equal(file2p$properties$name, folder1$properties$name)

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

    expect_silent(file3$delete(confirm=FALSE))
    expect_silent(folder2$delete(confirm=FALSE))
    expect_silent(file2$delete(confirm=FALSE))
    expect_silent(folder1$delete(confirm=FALSE))
    expect_silent(file1$delete(confirm=FALSE))
})


test_that("Methods work with filenames with special characters",
{
    test_name <- paste(make_name(5), "plus spaces and áccénts")
    src <- write_file(fname=file.path(tempdir(), test_name))

    expect_silent(od$upload_file(src, basename(src)))
    expect_silent(item <- od$get_item(basename(test_name)))
    expect_true(item$properties$name == basename(test_name))
    expect_silent(item$delete(confirm=FALSE))
})


test_that("Nested folder creation/deletion works",
{
    f1 <- make_name(10)
    f2 <- make_name(10)
    f3 <- make_name(10)

    it12 <- od$create_folder(file.path(f1, f2))
    expect_is(it12, "ms_drive_item")

    it1 <- od$get_item(f1)
    expect_is(it1, "ms_drive_item")

    replicate(30, it1$upload(write_file()))

    it123 <- it1$create_folder(file.path(f2, f3))
    expect_is(it123, "ms_drive_item")

    expect_silent(it1$delete(confirm=FALSE, by_item=TRUE))
})
