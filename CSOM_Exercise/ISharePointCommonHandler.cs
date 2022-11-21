using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

interface ISharePointCommonHandler
{
    public bool showAll(ClientContext context);

    public bool add(ClientContext context);

    public ListCollection getAll(ClientContext context);

}
