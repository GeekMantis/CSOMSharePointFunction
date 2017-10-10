#r "Microsoft.SharePoint.Client.Runtime.dll"
#r "Microsoft.SharePoint.Client.dll"

using System;
using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request.");

    // parse query parameter
    string siteCollection = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "SiteCollection", true) == 0)
        .Value;

     string action = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "Action", true) == 0)
        .Value;

     string groupName = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "GroupName", true) == 0)
        .Value;

     string user = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "UserName", true) == 0)
        .Value;

    string username = "";
    string password = "";

    // Get request body
    dynamic data = await req.Content.ReadAsAsync<object>();

    // Set name to query string or body data
    siteCollection = siteCollection ?? data?.SiteCollection;
    action = action ?? data?.Action;
    groupName = groupName ?? data?.GroupName;
    user = user ?? data?.UserName;

    if (siteCollection == null || action == null || groupName == null || user == null)
    {

        return req.CreateResponse(HttpStatusCode.BadRequest, "Please pass the following parameters: site collection, action, groupname and username on the query string or in the request body");
    }
    else 
    {
        System.Security.SecureString secureString = new System.Security.SecureString();
        foreach(char ch in password)
            secureString.AppendChar(ch);
    
        SharePointOnlineCredentials creds = new SharePointOnlineCredentials(username, secureString);
        
ClientContext client = new ClientContext(siteCollection);
client.Credentials = creds;
client.ExecuteQuery();
Web website = client.Web;
client.Load(website, w => w.AllProperties, w => w.SiteGroups, w => w.SiteUserInfoList, w => w.Webs,w => w.Title);
client.ExecuteQuery();
GroupCollection groupCollection = website.SiteGroups;
client.Load(groupCollection, groups => groups.Include(grps => grps.Users, grps => grps.Title));
client.ExecuteQuery();
User spuser;
foreach (Group group in groupCollection)
{
   if (group.Title.Equals(groupName))
   {
      UserCreationInformation userInfo = new UserCreationInformation();
      userInfo.LoginName = user;

             
      if (action.Equals("a"))
      {
         spuser = group.Users.Add(userInfo);
         group.Users.AddUser(spuser);
       }
       else if (action.Equals("r"))
       {
         spuser = group.Users.GetByLoginName(user);
         group.Users.Remove(spuser);
        }
        group.Update();
        website.Update();
        client.ExecuteQuery();
       
     }
  }


        

           
         
           

        
    }








        return req.CreateResponse(HttpStatusCode.OK, "OK");
    }

