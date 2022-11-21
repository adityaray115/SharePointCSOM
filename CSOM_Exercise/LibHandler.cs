using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.PortalLaunch;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

public class LibHandler : ISharePointCommonHandler
{
    public bool add(ClientContext context)
    {
        return true;
    }

    public bool showAll(ClientContext context)
    {
        _ = context ??
            throw new System.NullReferenceException("CSOM not initialized");
        try
        {
            ListCollection lists = getAll(context);
            if (lists.IsNullOrEmpty())
            {
                return false;
            }
            Console.WriteLine("Lists:");
            foreach (List list in lists)
            {
                context.Load(list.RootFolder);
                context.ExecuteQuery();
                Console.WriteLine("Name: {0}",list.Title);
                Console.Write("URL: {0}",context.Url.Split('/')[2] + list.RootFolder.ServerRelativeUrl);
                Console.WriteLine();
                Console.WriteLine();
            }
        }
        catch (Exception ex)
        {
            CSOMHandler.handleError(ex);
            return false;
        }
        return true;
    }

    public ListCollection getAll(ClientContext context)
    {
        try
        {
            ListHandler handler = new ListHandler();
            Web web = context.Web;
            context.Load(web.Lists, lists => lists.Where(lists => lists.Hidden == false && lists.IsSystemList == false && lists.BaseType==BaseType.DocumentLibrary));
            context.ExecuteQuery();
            return web.Lists;
        }
        catch (Exception ex)
        {
            CSOMHandler.handleError(ex);
            return null;
        }
    }

}
