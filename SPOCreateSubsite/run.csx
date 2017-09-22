#r "Microsoft.SharePoint.Client.Runtime.dll"
#r "Microsoft.SharePoint.Client.dll"

using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request.");

    
    // parse query parameter
    string siteUrl = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "Site", true) == 0).Value;
 
    log.Info(siteUrl);
    string userName = "";
    string password = "";

    //Subsite creation parameters 
    string Title = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "subSiteTitle", true) == 0)
        .Value;
         log.Info(Title);
    string SubSiteUrl = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "subSiteUrl", true) == 0)
        .Value;
    log.Info(SubSiteUrl);
    string Description = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "subSiteDescription", true) == 0)
        .Value;
        log.Info(Description);
    string ParentPermissions = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "ParentPermission", true) == 0)
        .Value;
       log.Info(ParentPermissions);
    string WebTemplate = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "WebTemplate", true) == 0)
        .Value;
     string Language = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "Language", true) == 0)
        .Value;
    

        
    // Get request body
    dynamic data = await req.Content.ReadAsAsync<object>();

    // Set name to query string or body data
   siteUrl = siteUrl ?? data?.Site;
   Title = Title ?? data?.subSiteTitle;
   SubSiteUrl = SubSiteUrl ?? data?.subSiteUrl;
   Description = Description ?? data?.subSiteDescription;
   ParentPermissions = ParentPermissions ?? data?.ParentPermission;
   WebTemplate = WebTemplate ?? data?.WebTemplate;
   Language = Language ?? data?.Language;

    System.Security.SecureString secureString = new System.Security.SecureString();
    foreach(char ch in password)
        secureString.AppendChar(ch);
    
    SharePointOnlineCredentials creds = new SharePointOnlineCredentials(userName, secureString);

    using (var clientContext = new ClientContext(siteUrl))
    {
        clientContext.Credentials = creds;
        WebCreationInformation wci = new WebCreationInformation();
        wci.Url = SubSiteUrl;
        wci.Title = Title;
        wci.Description = Description;
        wci.WebTemplate = WebTemplate;
        wci.UseSamePermissionsAsParentSite = Convert.ToBoolean(ParentPermissions);
        wci.Language = Convert.ToInt32(Language); 
       
        Web web = clientContext.Site.RootWeb.Webs.Add(wci);        
        clientContext.ExecuteQuery();
        return web == null
        ? req.CreateResponse(HttpStatusCode.BadRequest, "Error creating Subsite")
        : req.CreateResponse(HttpStatusCode.OK, "Site Created " + Title);

    }


   
    

    return null;
    
}
