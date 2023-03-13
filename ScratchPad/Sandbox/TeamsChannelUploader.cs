using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using System.Text;

namespace Sandbox
{
    public class TeamsChannelUploader
    {
        public async Task UploadFile()
        {
            // Can't use DefaultAzureCredential because that will use the VisualStudioCredential when a managed identity is not available.
            // An access token based on VisualStudioCredential does not include the scopes needed to upload to the Graph API.
            //
            // | Scope                          | VisualStudioCredential | AzureCliCredential |
            // |--------------------------------|------------------------|--------------------|
            // | Application.ReadWrite.All      | x                      |                    |
            // | AuditLog.Read.All              |                        | x                  |
            // | Directory.Read.All             | x                      |                    |
            // | Directory.AccessAsUser.All     |                        | x                  |
            // | email                          | x                      | x                  |
            // | Group.ReadWrite.All            |                        | x                  |
            // | IdentityUserFlow.ReadWrite.All | x                      |                    |
            // | openid                         | x                      | x                  |
            // | profile                        | x                      | x                  |
            // | User.Read                      | x                      |                    |
            // | User.ReadWrite.All             |                        | x                  |
            //
            // See GitHub issue: https://github.com/Azure/azure-sdk-for-net/issues/34843#issuecomment-1464108834
            var chainedCredential = new ChainedTokenCredential(
                new ManagedIdentityCredential(),
                new AzureCliCredential());

            // Before running with the Managed Identity make sure to add the permission Files.ReadWrite.All
            // See https://learn.microsoft.com/en-us/azure/app-service/scenario-secure-app-access-microsoft-graph-as-app?tabs=azure-powershell#grant-access-to-microsoft-graph

            // This variable is only here so we can inspect the token and verify scopes in https://jwt.io
            var chainedToken = chainedCredential.GetToken(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }));

            var graphServiceClient = new GraphServiceClient(
                chainedCredential,
                new[] { "https://graph.microsoft.com/.default" });

            // Retrieve the files folder of the Teams channel we want to upload files to.
            // Learn how to obtain the teams and channel id needed in the upload call:
            // https://www.c-sharpcorner.com/blogs/how-to-fetch-the-teams-id-and-channel-id-for-microsoft-teams
            var teamsId = "replaceWithTeamsId";
            var channelId = "replaceWithChannelId";

            var filesFolder = await graphServiceClient
                .Teams[teamsId]
                .Channels[channelId]
                .FilesFolder
                .GetAsync();

            // Create file content to upload
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(@"The contents of the file goes here - will be CSV eventually."));

            // Upload file
            await graphServiceClient
                .Drives[filesFolder!.ParentReference!.DriveId]
                .Items[filesFolder.Id]
                .ItemWithPath("uploadTest.txt")
                .Content
                .PutAsync(stream);
        }
    }
}