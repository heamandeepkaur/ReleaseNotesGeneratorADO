using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System.Linq;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using System.Text;
using Azure.Storage.Blobs;

namespace AdoReleaseNotes
{
    public static class ReleaseNotesWebhook
    {
        [FunctionName("ReleaseNotesWebhook")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            [Blob("releases/latest", FileAccess.Write, Connection = "StorageAccountConnectionStringDev")] Stream releaseNotes,
            [Blob("releases", Connection = "StorageAccountConnectionStringDev")] BlobContainerClient blobContainerClient,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            VssBasicCredential credentials = new VssBasicCredential(Environment.GetEnvironmentVariable("DevOps.Username"), Environment.GetEnvironmentVariable("DevOps.AccessToken"));
            VssConnection connection = new VssConnection(new Uri(Environment.GetEnvironmentVariable("DevOps.OrganizationURL")), credentials);

            //Time span of 14 days from today
            var dateSinceLastRelease = DateTime.Today.Subtract(new TimeSpan(14, 0, 0, 0));

            //Accumulate closed work items from the past 14 days in text format
            var workItems = GetClosedItems(connection, dateSinceLastRelease);
            var pulls = GetMergedPRs(connection, dateSinceLastRelease);


            string name = req.Query["name"];
            string releaseName = "Latest release";
            string releaseBody = "These are Work Items and Merged PR's for the last Sprint";

            //Extract data from request body
            if (req.ContentLength != null)
            {
                string requestBody = new StreamReader(req.Body).ReadToEnd();
                dynamic data = JsonConvert.DeserializeObject(requestBody);
                releaseName = data?.resource?.release?.name;
                releaseBody = data?.resource?.release?.description;
            }

            var responseMessage = String.Format("# {0} \n {1} \n\n" + "# Closed ADO Tickets:" + workItems + "\n\n# Changes Merged:" + pulls, releaseName, releaseBody);
            if (req.ContentLength != null)
            {
                var messageBytes = Encoding.UTF8.GetBytes(responseMessage);
                releaseNotes.Write(messageBytes, 0, messageBytes.Length);

                var blob = blobContainerClient.GetBlobClient(releaseName + ".md");
                var exists = await blob.ExistsAsync();
                if (!exists)
                {
                    var bytes = Encoding.UTF8.GetBytes(responseMessage);
                    MemoryStream stream = new(bytes);
                    await blob.UploadAsync(stream);
                }
            }
            responseMessage = Markdig.Markdown.ToHtml(responseMessage);
            return new OkObjectResult(responseMessage);
        }
        public static string GetClosedItems(VssConnection connection, DateTime releaseSpan)
        {
            string project = Environment.GetEnvironmentVariable("DevOps.ProjectName");
            string organizationURL = Environment.GetEnvironmentVariable("DevOps.OrganizationURL");
            var workItemTrackingHttpClient = connection.GetClient<WorkItemTrackingHttpClient>();

            //Query that grabs all of the Work Items marked "Done" in the last 14 days
            Wiql wiql = new Wiql()
            {
                Query = "Select [State], [Title] " +
                        "From WorkItems Where " +
                        "[System.TeamProject] = '" + project + "' " +
                        "And [System.State] = 'Closed' " +
                        "And [System.AreaPath] = 'RiskIQ\\EASM'" +
                        "And [Closed Date] >= '" + releaseSpan.ToString() + "' " +
                        "Order By [State] Asc, [Changed Date] Desc"
            };

            using (workItemTrackingHttpClient)
            {
                WorkItemQueryResult workItemQueryResult = workItemTrackingHttpClient.QueryByWiqlAsync(wiql).Result;

                if (workItemQueryResult.WorkItems.Count() != 0)
                {
                    List<int> list = new List<int>();
                    foreach (var item in workItemQueryResult.WorkItems)
                    {
                        list.Add(item.Id);
                    }

                    //Extraxt desired work item fields
                    string[] fields = { "System.Id", "System.Title" };
                    var workItems = workItemTrackingHttpClient.GetWorkItemsAsync(list, fields, workItemQueryResult.AsOf).Result;

                    //Format Work Item info into text
                    string txtWorkItems = string.Empty;
                    int counter = 1;
                    foreach (var workItem in workItems)
                    {
                        //LINK FORMAT - https://dev.azure.com/{ORG}/{PROJECT}/_workitems/edit/{WORKITEMNUMBER}/
                        string workItemLink = String.Format("{0}{1}/_workitems/edit/{2}", organizationURL, project, workItem.Id);
                        string link = String.Format("[{0}]({1})", workItem.Id, workItemLink);
                        txtWorkItems += String.Format("\n {0}. {1} - {2}", counter, link, workItem.Fields["System.Title"]);
                        //txtWorkItems += "\n " + counter + ". #[" + workItem.Id + "](" + organizationURL + project + "/_workitems/edit/" + workItem.Id + ") - " + workItem.Fields["System.Title"];
                        counter++;
                    }
                    return txtWorkItems;
                }
                return string.Empty;
            }
        }

        public static string GetMergedPRs(VssConnection connection, DateTime releaseSpan)
        {
            string projectName = Environment.GetEnvironmentVariable("DevOps.ProjectName");
            string organizationURL = Environment.GetEnvironmentVariable("DevOps.OrganizationURL");
            string repoName = Environment.GetEnvironmentVariable("DevOps.RepoName");
            var gitClient = connection.GetClient<GitHttpClient>();

            using (gitClient)
            {
                //Get first repo in project
                var releaseRepos = gitClient.GetRepositoriesAsync().Result;
                var releaseRepo = releaseRepos.Find(repo => repo.Name == repoName);

                //Grabs all completed PRs merged into master branch
                List<GitPullRequest> prs = gitClient.GetPullRequestsAsync(
                   releaseRepo.Id,
                   new GitPullRequestSearchCriteria()
                   {
                       TargetRefName = "refs/heads/main",
                       Status = PullRequestStatus.Completed

                   }).Result;

                if (prs.Count != 0)
                {
                    //Query that grabs the merged PR's since the specified date
                    var pulls = from p in prs
                                where p.ClosedDate >= releaseSpan
                                select p;

                    //Format PR info into text
                    var txtPRs = string.Empty;
                    int counter = 1;
                    foreach (var pull in pulls)
                    {
                        // link format - https://dev.azure.com/{org}/{project}/_git/{repo}/pullrequest/{pr id}
                        string prLink = organizationURL + projectName + "/_git/" + repoName + "/pullrequest/" + pull.PullRequestId;
                        string link = String.Format("[{0}]({1})", pull.PullRequestId, prLink);
                        txtPRs += String.Format("\n {0}. #{1}-{2}", counter, link, pull.Title);

                        counter++;
                    }

                    return txtPRs;
                }
                return string.Empty;
            }
        }
    }
}

