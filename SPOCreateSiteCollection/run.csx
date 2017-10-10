#r "Microsoft.SharePoint.Client.Runtime.dll"
#r "Microsoft.SharePoint.Client.dll"
#r "Microsoft.Online.SharePoint.Client.Tenant.dll"

using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request.");

    
  
    

    //SiteCollection creation parameters 
    string mainSiteCollection = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "mainSiteCollection", true) == 0)
        .Value;
    string newSiteCollectionTitle = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "newSiteCollectionTitle", true) == 0)
        .Value;         
    string newSiteCollectionUrl = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "newSiteCollectionUrl", true) == 0)
        .Value;
    
    string newSiteCollectionDescription = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "newSiteCollectionDescription", true) == 0)
        .Value;
        
    string newSiteCollectionOwner = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "newSiteCollectionOwner", true) == 0)
        .Value;
       
    string newSiteCollectionWebTemplate = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "newSiteCollectionWebTemplate", true) == 0)
        .Value;
    string newSiteCollectionStorageMaximumLevel = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "newSiteCollectionStorageMaximumLevel", true) == 0)
        .Value;
    string newSiteCollectionUserCodeMaximumLevel = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "newSiteCollectionUserCodeMaximumLevel", true) == 0)
        .Value;
    

        
    // Get request body
    dynamic data = await req.Content.ReadAsAsync<object>();

    // Set name to query string or body data
   mainSiteCollection = mainSiteCollection ?? data?.mainSiteCollection;
   newSiteCollectionTitle = newSiteCollectionTitle ?? data?.newSiteCollectionTitle;
   newSiteCollectionUrl = newSiteCollectionUrl ?? data?.newSiteCollectionUrl;
   newSiteCollectionDescription = newSiteCollectionDescription ?? data?.newSiteCollectionDescription;
   newSiteCollectionOwner = newSiteCollectionOwner ?? data?.newSiteCollectionOwner;
   newSiteCollectionWebTemplate = newSiteCollectionWebTemplate ?? data?.newSiteCollectionWebTemplate;
   newSiteCollectionStorageMaximumLevel = newSiteCollectionStorageMaximumLevel ?? data?.newSiteCollectionStorageMaximumLevel;
   newSiteCollectionUserCodeMaximumLevel = newSiteCollectionUserCodeMaximumLevel ?? data?.newSiteCollectionUserCodeMaximumLevel;

   

            //Open the Tenant Administration Context with the Tenant Admin Url
            using (ClientContext tenantContext = new ClientContext(mainSiteCollection))
            {
                //Authenticate with a Tenant Administrator
                string userName = "";
                string password = "";
                 System.Security.SecureString secureString = new System.Security.SecureString();
                foreach(char ch in password)
                {
                    secureString.AppendChar(ch);
                }
                    
    
                SharePointOnlineCredentials creds = new SharePointOnlineCredentials(userName, secureString);

           
                tenantContext.Credentials = creds;

                var tenant = new Microsoft.Online.SharePoint.TenantAdministration.Tenant(tenantContext);

                //Properties of the New SiteCollection
                var siteCreationProperties = new Microsoft.Online.SharePoint.TenantAdministration.SiteCreationProperties();
                
                //New SiteCollection Url
                siteCreationProperties.Url = newSiteCollectionUrl;
                
                //Title of the Root Site
                siteCreationProperties.Title = newSiteCollectionTitle;

                //Login name of Owner
                siteCreationProperties.Owner = newSiteCollectionOwner;
                
                //Template of the Root Site. Using Team Site for now.
                siteCreationProperties.Template = newSiteCollectionWebTemplate;

                //Storage Limit in MB
                siteCreationProperties.StorageMaximumLevel = Convert.ToInt32(newSiteCollectionStorageMaximumLevel);

                //UserCode Resource Points Allowed
                siteCreationProperties.UserCodeMaximumLevel = Convert.ToInt32(newSiteCollectionUserCodeMaximumLevel);
               
                //Create the SiteCollection
                Microsoft.Online.SharePoint.TenantAdministration.SpoOperation spo = tenant.CreateSite(siteCreationProperties);

                tenantContext.Load(tenant);
                
                //We will need the IsComplete property to check if the provisioning of the Site Collection is complete.
                tenantContext.Load(spo, i => i.IsComplete);
                
                tenantContext.ExecuteQuery();

                //Check if provisioning of the SiteCollection is complete.
                while (!spo.IsComplete)
                {
                    //Wait for 30 seconds and then try again
                    System.Threading.Thread.Sleep(30000);
                    spo.RefreshLoad();
                    tenantContext.ExecuteQuery();
                }

               return spo.IsComplete == false
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Error creating Site Collection")
                : req.CreateResponse(HttpStatusCode.OK, "Site Created " + newSiteCollectionUrl);

            }



   
    

    return null;
    
}
