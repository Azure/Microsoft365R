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

test_that("OneDrive personal works",
{
    expect_is(od, "ms_drive")

    ls <- od$list_items()
    expect_is(ls, "data.frame")

    newfolder <- make_name()
    expect_silent(od$create_folder(newfolder))

    src <- "../resources/file.json"
    dest <- file.path(newfolder, basename(src))
    newsrc <- tempfile()
    expect_silent(od$upload_file(src, dest=dest))
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

    expect_silent(od$delete_item(newfolder, confirm=FALSE))
})


test_that("Drive item methods work",
{
    root <- od$get_item("/")
    expect_is(root, "ms_drive_item")
    expect_equal(root$properties$name, "root")
    expect_equal(root$get_path(), "/")

    rootp <- root$get_parent_folder()
    expect_is(rootp, "ms_drive_item")
    expect_equal(rootp$properties$name, "root")

    tmpname1 <- make_name(10)
    folder1 <- root$create_folder(tmpname1)
    expect_is(folder1, "ms_drive_item")
    expect_true(folder1$is_folder())
    expect_equal(folder1$get_path(), paste0("/", tmpname1))

    folder1p <- folder1$get_parent_folder()
    expect_equal(rootp$properties$name, "root")

    tmpname2 <- make_name(10)
    folder2 <- folder1$create_folder(tmpname2)
    expect_is(folder2, "ms_drive_item")
    expect_true(folder2$is_folder())
    expect_equal(folder2$get_path(), paste0("/", tmpname1, "/", tmpname2))

    folder2p <- folder2$get_parent_folder()
    expect_equal(folder2p$properties$name, folder1$properties$name)

    src <- write_file()
    expect_silent(file1 <- root$upload(src))
    expect_is(file1, "ms_drive_item")
    expect_false(file1$is_folder())
    expect_error(file1$create_folder("bad"))
    expect_equal(file1$get_path(), paste0("/", basename(src)))

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
    expect_equal(file2$get_path(), paste0("/", tmpname1, "/", basename(src)))

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

    expect_silent(file1$delete(confirm=FALSE))
    expect_silent(folder1$delete(confirm=FALSE))
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


test_that("Get item by ID works",
{
    dir1 <- make_name(10)
    src <- "../resources/file.json"

    obj <- od$upload_file(src, file.path(dir1, "file.json"))
    id <- obj$properties$id
    obj2 <- od$get_item(itemid=id)
    expect_identical(obj$properties$id, obj2$properties$id)

    expect_silent(obj2$delete(confirm=FALSE))
})


test_that("Folder upload/download works",
{
    srcdir <- tempfile()
    srcdir2 <- tempfile(tmpdir=srcdir)
    dir.create(srcdir2, showWarnings=FALSE, recursive=TRUE)
    src1 <- write_file(srcdir)
    src2 <- write_file(srcdir)
    src3 <- write_file(srcdir2)
    src4 <- write_file(srcdir2)

    root <- od$get_item("/")

    # serial, not recursive
    destdir1 <- basename(tempfile())
    returndir1 <- tempfile()

    expect_silent(root$upload(srcdir, destdir1, recursive=FALSE, parallel=FALSE))

    obj1 <- root$get_item(destdir1)
    expect_silent(obj1$download(returndir1, recursive=FALSE, parallel=FALSE))
    expect_silent(obj1$delete(confirm=FALSE))

    files_identical(src1, file.path(returndir1, basename(src1)))
    files_identical(src2, file.path(returndir1, basename(src2)))
    expect_false(file.exists(file.path(returndir1, basename(src3))))
    expect_false(file.exists(file.path(returndir1, basename(src4))))

    # serial, recursive
    destdir2 <- basename(tempfile())
    returndir2 <- tempfile()

    expect_silent(root$upload(srcdir, destdir2, recursive=TRUE, parallel=FALSE))

    obj2 <- root$get_item(destdir2)
    expect_silent(obj2$download(returndir2, recursive=TRUE, parallel=FALSE))
    expect_silent(obj2$delete(confirm=FALSE))

    files_identical(src1, file.path(returndir2, basename(src1)))
    files_identical(src2, file.path(returndir2, basename(src2)))
    files_identical(src3, file.path(returndir2, basename(srcdir2), basename(src3)))
    files_identical(src4, file.path(returndir2, basename(srcdir2), basename(src4)))

    # parallel, not recursive
    destdir3 <- basename(tempfile())
    returndir3 <- tempfile()

    expect_silent(root$upload(srcdir, destdir3, recursive=FALSE, parallel=TRUE))

    obj3 <- root$get_item(destdir3)
    expect_silent(obj3$download(returndir3, recursive=FALSE, parallel=TRUE))
    expect_silent(obj3$delete(confirm=FALSE))

    files_identical(src1, file.path(returndir3, basename(src1)))
    files_identical(src2, file.path(returndir3, basename(src2)))
    expect_false(file.exists(file.path(returndir3, basename(src3))))
    expect_false(file.exists(file.path(returndir3, basename(src4))))

    # parallel, recursive
    destdir4 <- basename(tempfile())
    returndir4 <- tempfile()

    expect_silent(root$upload(srcdir, destdir4, recursive=TRUE, parallel=TRUE))

    obj4 <- root$get_item(destdir4)
    expect_silent(obj4$download(returndir4, recursive=TRUE, parallel=TRUE))
    expect_silent(obj4$delete(confirm=FALSE))

    files_identical(src1, file.path(returndir4, basename(src1)))
    files_identical(src2, file.path(returndir4, basename(src2)))
    files_identical(src3, file.path(returndir4, basename(srcdir2), basename(src3)))
    files_identical(src4, file.path(returndir4, basename(srcdir2), basename(src4)))
})


teardown({
    options(opt_use_itemid)
})
