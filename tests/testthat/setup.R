make_name <- function(n=20)
{
    paste0(sample(letters, n, TRUE), collapse="")
}

write_file <- function(dir=tempdir(), size=1000)
{
    fname <- tempfile(tmpdir=dir)
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
