tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
site_name <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_NAME")
site_url <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_URL")
site_id <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_ID")

if(tenant == "" || app == "" || site_name == "" || site_url == "" || site_id == "")
    skip("SharePoint tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("SharePoint authentication tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Group.ReadWrite.All", "Directory.Read.All",
                                     "Sites.ReadWrite.All", "Sites.Manage.All"))
if(is.null(tok))
    skip("SharePoint authentication tests skipped: no access to tenant")

site <- try(call_graph_endpoint(tok, file.path("sites", site_id)), silent=TRUE)
if(inherits(site, "try-error"))
    skip("SharePoint authentication tests skipped: service not available")

test_that("Sharepoint authentication works",
{
    site <- get_sharepoint_site(site_id=site_id, token=tok)
    expect_is(site, "ms_site")

    sitelist <- list_sharepoint_sites(token=tok)
    expect_is(sitelist, "list")
    expect_true(all(sapply(sitelist, inherits, "ms_site")))

    sitelist2 <- list_sharepoint_sites(tenant=tenant, app=app)
    expect_is(sitelist2, "list")
    expect_true(all(sapply(sitelist2, inherits, "ms_site")))
    expect_true(all(mapply(
        function(s1, s2) identical(s1$properties$id, s2$properties$id,
        sitelist1,
        sitelist2
    ))))

    site2 <- get_sharepoint_site(site_name=site_name, token=tok)
    expect_is(site2, "ms_site")
    expect_identical(site$properties$id, site2$properties$id)

    site3 <- get_sharepoint_site(site_url=site_url, token=tok)
    expect_is(site3, "ms_site")
    expect_identical(site$properties$id, site3$properties$id)

    site4 <- get_sharepoint_site(site_id=site_id, tenant=tenant, app=app)
    expect_is(site4, "ms_site")
    expect_identical(site$properties$id, site4$properties$id)

    site5 <- get_sharepoint_site(site_url=site_url, tenant=tenant, app=app)
    expect_is(site5, "ms_site")
    expect_identical(site$properties$id, site5$properties$id)

    site6 <- get_sharepoint_site(site_name=site_name, tenant=tenant, app=app)
    expect_is(site6, "ms_site")
    expect_identical(site$properties$id, site6$properties$id)
})
