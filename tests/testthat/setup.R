make_name <- function(n=20)
{
    paste0(sample(letters, n, TRUE), collapse="")
}

write_file <- function(dir=tempdir(), size=1000, fname=tempfile(tmpdir=dir))
{
    bytes <- openssl::rand_bytes(size)
    writeBin(bytes, fname)
    fname
}

files_identical <- function(set1, set2)
{
    all(mapply(function(f1, f2)
    {
        s1 <- file.size(f1)
        s2 <- file.size(f2)
        s1 == s2 && identical(readBin(f1, "raw", s1), readBin(f2, "raw", s2))
    }, set1, set2))
}

filter_esc <- function(x)
{
    gsub("'", "''", x)
}

get_test_token <- function(tenant, app, scopes, ...)
{
    # if using MS365 CLI or Azure CLI app IDs...
    # - with org tenant: set to .default scope
    # - with consumers tenant: fail
    consumers_tenant <- tenant %in% c("consumers", "9188040d-6c67-4c5b-b112-36a304b66dad")
    special_app <- app %in% c("31359c7f-bd7e-475c-86db-fdb8c937548e", "04b07795-8ddb-461a-bbee-02f9e1bf7b46")
    if(special_app)
    {
        if(consumers_tenant)
            return(NULL)
        else scopes <- ".default"
    }

    scopes <- c(file.path("https://graph.microsoft.com", scopes), "openid", "offline_access")
    tok <- try(AzureAuth::get_azure_token(scopes, tenant, app, ..., version=2), silent=TRUE)
    if(inherits(tok, "try-error"))
        return(NULL)
    tok
}

