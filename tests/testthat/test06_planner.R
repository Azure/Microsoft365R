tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
site_url <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_URL")
site_id <- Sys.getenv("AZ_TEST_SHAREPOINT_SITE_ID")

if(tenant == "" || app == "" || site_url == "" || site_id == "")
    skip("SharePoint tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Planner tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Group.ReadWrite.All", "Directory.Read.All",
                                     "Sites.ReadWrite.All", "Sites.Manage.All"))
if(is.null(tok))
    skip("Planner tests skipped: no access to tenant")

site <- try(call_graph_endpoint(tok, file.path("sites", site_id)), silent=TRUE)
if(inherits(site, "try-error"))
    skip("Planner tests skipped: service not available")

site <- ms_site$new(tok, tenant, site)

test_that("Planner methods work",
{
    expect_is(site, "ms_site")

    grp <- site$get_group()

    plans <- grp$list_plans()
    expect_is(plans, "list")
    expect_true(all(sapply(plans, inherits, "ms_plan")))

    plan_title <- plans[[1]]$properties$title

    # filtering not yet implemented
    expect_error(lstpager <- grp$list_plans(filter=sprintf("title eq '%s'", filter_esc(plan_title)), n=NULL))
    # expect_is(lstpager, "ms_graph_pager")
    # plan0 <- lstpager$value
    # expect_true(length(plan0) == 1 && inherits(plan0[[1]], "ms_plan"))

    plan1 <- grp$get_plan(plan_title=plan_title)
    expect_is(plan1, "ms_plan")

    bkts <- plan1$list_buckets()
    expect_is(bkts, "list")
    expect_true(all(sapply(bkts, inherits, "ms_plan_bucket")))

    tasks <- plan1$list_tasks()
    expect_is(tasks, "list")
    expect_true(all(sapply(tasks, inherits, "ms_plan_task")))
})
