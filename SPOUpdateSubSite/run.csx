#r "Microsoft.SharePoint.Client.Runtime.dll"
#r "Microsoft.SharePoint.Client.dll"

using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request. Update");   

    // parse subsite query parameter
    string siteUrl = req.GetQueryNameValuePairs()
    .FirstOrDefault(q => string.Compare(q.Key, "Site", true) == 0)
    .Value;
  
    
    string userName = "";
    string password = "";

    //Subsite Update parameters 
    string Title = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "SubSiteTitle", true) == 0)
        .Value;
         
    string SubSiteUrl = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "SubSiteUrl", true) == 0)
        .Value;
    
    string Description = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "SubSiteDescription", true) == 0)
        .Value;
       
    string InheritPermissions = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "InheritPermissions", true) == 0)
        .Value;

         string CopyRoleAssignments = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "CopyRoleAssignments", true) == 0)
        .Value;

         string ClearUniquePermissions = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "ClearUniquePermissions", true) == 0)
        .Value;
     

       string InheritNavigation = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "InheritNavigation", true) == 0)
        .Value;
       
    
   

    // Get request body
    dynamic data = await req.Content.ReadAsAsync<object>();
  

   // Set name to query string or body data
   siteUrl = siteUrl ?? data?.Site;
   Title = Title ?? data?.SubSiteTitle;
   SubSiteUrl = SubSiteUrl ?? data?.SubSiteUrl;
   Description = Description ?? data?.SubSiteDescription;
   InheritPermissions = InheritPermissions ?? data?.InheritPermissions;
   CopyRoleAssignments = CopyRoleAssignments ?? data?.CopyRoleAssignments;
   ClearUniquePermissions = ClearUniquePermissions ?? data?.ClearUniquePermissions;
   InheritNavigation = InheritNavigation ?? data?.InheritNavigation;
   log.Info(siteUrl);
   log.Info(InheritNavigation);
   log.Info(Description);
   log.Info(SubSiteUrl);
   log.Info("TITLE:" + Title);
   log.Info("Permissions " + InheritPermissions);

    //Authenticate to SharePoint with credentials
    System.Security.SecureString secureString = new System.Security.SecureString();
    //Create a Secure String
    foreach(char ch in password)
    {
        secureString.AppendChar(ch);
    }    
    
    SharePointOnlineCredentials creds = new SharePointOnlineCredentials(userName, secureString);

     using (var clientContext = new ClientContext(siteUrl))
    {
        clientContext.Credentials = creds;
        Web web = clientContext.Web;
        bool copyRole = false;
        bool cleanUnique = false;

        if(InheritPermissions.Equals("false"))
        {
            log.Info("Breaking Inheritance");
            if(CopyRoleAssignments.Equals("true"))
            copyRole = true;

            if(ClearUniquePermissions.Equals("true"))
            cleanUnique = true;

            web.BreakRoleInheritance(copyRole,  cleanUnique);
        }
        else if(InheritPermissions.Equals("true"))
        {
            log.Info("Reseting role inheritance");
            web.ResetRoleInheritance();
        }
        


        clientContext.Load(web);
        clientContext.ExecuteQuery();
        
        if(!Title.Equals(""))
        web.Title = Title;

        if(!Description.Equals(""))
        web.Description = Description;

        if(!SubSiteUrl.Equals(""))
        web.ServerRelativeUrl = SubSiteUrl;

        if(!InheritNavigation.Equals(""))
        web.Navigation.UseShared = Convert.ToBoolean(InheritNavigation);

      
      
        web.Update();        
        
             
        return web == null
        ? req.CreateResponse(HttpStatusCode.BadRequest, "Error Updating Subsite")
        : req.CreateResponse(HttpStatusCode.OK, "Subsite Updated " + SubSiteUrl);
    
    }     



   return null;
}
