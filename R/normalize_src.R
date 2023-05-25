normalize_src <- function(src)
{
    UseMethod("normalize_src")
}


normalize_src.character <- function(src)
{
    con <- file(src, open="rb")
    size <- file.size(src)
    list(con=con, size=size)
}


normalize_src.textConnection <- function(src)
{
    # convert to raw connection
    src <- charToRaw(paste0(readLines(src), collapse="\n"))
    size <- length(src)
    con <- rawConnection(src)
    list(con=con, size=size)
}


normalize_src.rawConnection <- function(src)
{
    # need to read the data to get object size (!)
    size <- 0
    repeat
    {
        x <- readBin(src, "raw", n=1e6)
        if(length(x) == 0)
            break
        size <- size + length(x)
    }
    seek(src, 0) # reposition connection after reading
    list(con=src, size=size)
}

