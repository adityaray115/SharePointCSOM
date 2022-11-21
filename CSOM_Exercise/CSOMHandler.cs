using Microsoft.IdentityModel.Tokens;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.PortalLaunch;
using Microsoft.SharePoint.News.DataModel;
using System; 
using System.Collections.Generic; 
using System.Linq; 
using System.Net;
using System.Reflection.Metadata.Ecma335;
using System.Reflection.PortableExecutable;
using System.Runtime.ExceptionServices;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

public class CSOMHandler
{
    public static ClientContext context;

    public static Uri siteUri;
    public static string ID;
    public static SecureString securePass;
    public static bool connetToSharePoint()
    {
        try
        {
            Console.WriteLine("Enter Site:");
            string URL = Console.ReadLine();

            while (URL.IsNullOrEmpty())
            {
                Console.WriteLine("Site Name can't be empty");
                Console.WriteLine("Enter Site:");
                URL = Console.ReadLine();
            }

            siteUri = new Uri(URL!);


            Console.WriteLine("Enter UserID");
            string ID = Console.ReadLine();
            while (ID.IsNullOrEmpty())
            {
                Console.WriteLine("ID can't be empty");
                Console.WriteLine("Enter ID");
                ID = Console.ReadLine();
            }

            Console.WriteLine("Enter Password");
            string pass = Console.ReadLine();
            while (pass.IsNullOrEmpty())
            {
                Console.WriteLine("Password can't be empty");
                Console.WriteLine("Enter Password");
                pass = Console.ReadLine();
            }
            securePass = new NetworkCredential("", pass).SecurePassword;
            AuthenticationManager manager = new AuthenticationManager();
            context = manager.GetContext(siteUri, ID, securePass);
            Console.WriteLine(context);

        }
        catch (Exception ex)
        {

            if (ex is UriFormatException)
            {
                Console.WriteLine("Invalid URL");
            }
            else
                Console.WriteLine(ex.ToString());
            return false;
        }
        return true;
    }

    public static void dummy_func()
    {

    }

    public static bool showAllListsAndLibraries()
    {
        try
        {
            ISharePointCommonHandler handler;

            bool success = true;
            handler = new ListHandler();
            success = handler.showAll(context);

            if (success == false)
            {
                return false;
            }

            handler = new LibHandler();
            success = handler.showAll(context);

            if (success == false)
            {
                return false;
            }
        }
        catch (Exception ex) {
            handleError(ex);
            return false;
        }
        return true;

    }

    internal static void handleError(Exception ex)
    {
        if (ex is WebException)
        {
            var ex2 = ex as WebException;
            var response = ex2.Response as HttpWebResponse;

            if (response != null)
            {

                if (response.StatusCode == HttpStatusCode.NotFound)
                    Console.WriteLine("Site Not Found");
                else if (response.StatusCode == HttpStatusCode.Forbidden)
                    Console.WriteLine("Permission Denied");
                else Console.WriteLine(response.StatusCode);
            }
        }
        else if (ex is KeyNotFoundException)
        {
            Console.WriteLine("Credentials Are wrong");
        }
        else if (ex is ServerException)
        {
            Console.WriteLine("List/Library doesn't Exist at the given Site");
        }
        else
        {
            Console.WriteLine(ex.GetBaseException());
        }
    }

    public static bool listWithMoreThanN(int N) {
        try
        {
            ISharePointCommonHandler handler;

            handler = new ListHandler();
            ListCollection lists = handler.getAll(context);

            var lists_updated = lists.Where(item => item.ItemCount >= N);

            Console.WriteLine("");
            Console.WriteLine("Lists with more than {0} items:", N);
            foreach (var item in lists_updated)
            {
                Console.WriteLine(item.Title);
            }


            handler = new LibHandler();
            ListCollection libs = handler.getAll(context);

            var libs_updated = lists.Where(item => item.ItemCount >= N);

            Console.WriteLine("");
            Console.WriteLine("Libs with more than {0} items:", N);
            foreach (var item in lists_updated)
            {
                Console.WriteLine(item.Title);
            }
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            handleError(ex);
            return false;
        }
        return true;
    }

    internal static bool SelectiveSearch()
    {
        try
        {
            ISharePointCommonHandler handler;

            Console.WriteLine("Select Option:");
            Console.WriteLine("1: List");
            Console.WriteLine("2: Library");
            int choice = Convert.ToInt32(Console.ReadLine());
            if (choice == 1)
            {
                handler = new ListHandler();
                Console.Write("Enter List Name:  ");
            }
            else if (choice == 2)
            {
                handler = new LibHandler();
                Console.Write("Enter Lib Name:   ");
            }
            else
            {
                Console.WriteLine("Invalid Option");
                return true;
            }

            string name = Console.ReadLine();

            if (name.IsNullOrEmpty())
            {
                Console.WriteLine("Empty Name Entered");
                return true;
            }


            Console.Write("Enter No. of Cols");
            int numOfCols = Convert.ToInt32(Console.ReadLine());
            HashSet<string> cols = new HashSet<string>();
            for (int i = 0; i < numOfCols; i++)
            {
                Console.Write("Col Name:  ");
                string colName = Console.ReadLine();
                while (colName.IsNullOrEmpty())
                    colName = Console.ReadLine();
                cols.Add(colName);
            }

            List list = context.Web.Lists.GetByTitle(name);

            context.Load(list.Fields);
            context.Load(list.RootFolder);

            CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();

            ListItemCollection oCollection = list.GetItems(camlQuery);
            context.Load(oCollection);
            context.ExecuteQuery();

            Dictionary<string, string> name_dict = new Dictionary<string, string>();

            foreach (var field in list.Fields)
            {
                if (field.Title == "Title")
                    Console.WriteLine(field.InternalName + " " + field.StaticName);
                if (cols.Contains(field.Title))
                {
                    name_dict[field.Title] = field.InternalName;
                }
            }

            if (choice == 1)
            {
                if (cols.Contains("Title"))
                    name_dict["Title"] = "Title";
            }
            if (choice == 2)
                if (cols.Contains("Name"))
                    name_dict["Name"] = "FileLeafRef";

            if (name_dict.Count != numOfCols)
            {
                Console.WriteLine("Not all Cols Found in the List");
                return true;
            }


            System.Xml.Serialization.XmlSerializer xmlSerializer = new System.Xml.Serialization.XmlSerializer(typeof(List<ListInfo>));
            StreamReader reader = new StreamReader(@"C:\\Result.xml");
            List<ListInfo> file;
            XmlReader rd = new XmlTextReader(@"C:\\Result.xml");
            Console.WriteLine(rd);
            if(chk())
            {
                file = (List<ListInfo>)xmlSerializer.Deserialize(reader);
            }
            else
            {
                file = new List<ListInfo>();
            }

            reader.Close();

            var lists = from ls in file where ls.name == name select ls;

            if(lists.Count()==0)
            {
                ListInfo xmlobj=new ListInfo();
                xmlobj.name = name;
                xmlobj.url = (new Uri(context.Url)).GetLeftPart(UriPartial.Authority);
                xmlobj.type = choice == 1 ? "Generic List" : "Document Library";
                xmlobj.item = new List<Item>();

                foreach (ListItem item in oCollection)
                {
                    var row = new Item();
                    row.name = choice == 1 ? item["Title"] == null ? "" : item["Title"].ToString() : item["FileLeafRef"] == null ? "" : item["FileLeafRef"].ToString();
                    row.ID = item["ID"].ToString();
                    row.Url = context.Url.Split('/')[2] + list.RootFolder.ServerRelativeUrl;

                    row.Column = new List<column>();
                    foreach (var col_name in cols)
                    {
                        row.Column.Add(new column
                        {
                            name = col_name,
                            value = item[name_dict[col_name]] == null ? "" : item[name_dict[col_name]].ToString()
                        });
                    }
                    xmlobj.item.Add(row);
                }
                file.Add(xmlobj);
            }
            else
            {
                var lst = lists.First(); 

                foreach (ListItem item in oCollection)
                {
                    if((from row in lst.item where row.ID == item["ID"].ToString() select row).Count()>0 )
                    {
                        var row = (from rw in lst.item where rw.ID == item["ID"].ToString() select rw).First();

                        foreach (var col_name in cols)
                        {
                            if((from col in row.Column where col.name == col_name select col).Count()==0)
                            {
                                row.Column.Add(new column
                                {
                                    name = col_name,
                                    value = item[name_dict[col_name]] == null ? "" : item[name_dict[col_name]].ToString()
                                });
                            }
                            else
                            {
                                (from col in row.Column where col.name == col_name select col).First().value = item[name_dict[col_name]] == null ? "" : item[name_dict[col_name]].ToString();
                            }
                        }
                    }
                    else
                    {
                        var row = new Item();
                        row.name = choice == 1 ? item["Title"] == null ? "" : item["Title"].ToString() : item["FileLeafRef"]==null?"":item["FileLeafRef"].ToString();
                        row.Url = context.Url.Split('/')[2] + list.RootFolder.ServerRelativeUrl;
                        row.ID = item["ID"].ToString();
                        row.Column = new List<column>();
                        foreach (var col_name in cols)
                        {
                            row.Column.Add(new column
                            {
                                name = col_name,
                                value = item[name_dict[col_name]] == null ? "" : item[name_dict[col_name]].ToString()
                            });
                        }
                        lst.item.Add(row);

                    }
                }
            }
                

            var writer = new StreamWriter(@"C:\\Result.xml");
            
            xmlSerializer.Serialize(writer, file);
            writer.Close();
        }
        catch (Exception ex)
        {
            if (ex is Microsoft.SharePoint.Client.PropertyOrFieldNotInitializedException)
            {
                Console.WriteLine("Invalid Field Entry Detected");
                return false;
            }
            handleError(ex);
            return false;
        }
        return true;
    }

    public static bool chk()
    {
        System.Xml.Serialization.XmlSerializer xmlSerializer = new System.Xml.Serialization.XmlSerializer(typeof(List<ListInfo>));
        StreamReader reader = new StreamReader(@"C:\\Result.xml");
        List<ListInfo> file;
        XmlReader rd = new XmlTextReader(@"C:\\Result.xml");
        Console.WriteLine(rd);

        try
        {
            file = (List<ListInfo>)xmlSerializer.Deserialize(reader);
            reader.Close();
            return true;
        }
        catch (System.Exception ex)
        {
            reader.Close();
            return false;
            //if(ex is )
        }
    }

    public static bool addMItems()
    {
        try
        {
            Console.Write("Enter No. of Rows");
            int numOfRows = Convert.ToInt32(Console.ReadLine());

            ISharePointCommonHandler handler = new ListHandler();

            var lists = handler.getAll(context);

            Dictionary<FieldType, Type> _fieldTypes = new Dictionary<FieldType, Type>()
        {
            { FieldType.Guid, typeof(Guid) },
            { FieldType.Attachments, typeof(bool)},
            {FieldType.Boolean, typeof(bool)},
            {FieldType.Choice, typeof(string)},
            {FieldType.CrossProjectLink, typeof(bool)},
            {FieldType.DateTime, typeof(DateTime)},
            {FieldType.Lookup, typeof(FieldLookupValue)},
            {FieldType.ModStat, typeof(int)},
            {FieldType.MultiChoice, typeof(string[])},
            {FieldType.Number, typeof(double) },
            {FieldType.Recurrence, typeof(bool)},
            {FieldType.Text, typeof(string)},
            {FieldType.URL, typeof(FieldUrlValue)},
            {FieldType.User, typeof(FieldUserValue)},
            {FieldType.WorkflowStatus, typeof(int)},
            {FieldType.ContentTypeId, typeof(ContentTypeId)},
            {FieldType.Note, typeof(string)},
            {FieldType.Counter, typeof(int)},
            {FieldType.Computed, typeof(string)},
            {FieldType.Integer, typeof(int)},
            {FieldType.File, typeof(string)}
        };

            HashSet<FieldType> _ints = new HashSet<FieldType> { FieldType.ModStat, FieldType.WorkflowStatus, FieldType.Counter, FieldType.Integer };

            HashSet<FieldType> _bools = new HashSet<FieldType> { FieldType.Attachments, FieldType.Boolean, FieldType.CrossProjectLink, FieldType.Recurrence };

            HashSet<FieldType> _strings = new HashSet<FieldType> { FieldType.Choice, FieldType.Text, FieldType.Note, FieldType.Computed, FieldType.File };

            HashSet<FieldType> _datetime = new HashSet<FieldType> { FieldType.DateTime };

            HashSet<FieldType> _double = new HashSet<FieldType> { FieldType.DateTime };


            Object obj = new Object();

            Parallel.ForEach(lists, lst=>
            {
                AuthenticationManager manager = new AuthenticationManager();
                ClientContext context2 = manager.GetContext(siteUri, ID, securePass);
                var list = context2.Web.Lists.GetById(lst.Id);


                var Fields = list.Fields;
                context2.Load(Fields);
                context2.ExecuteQuery();
                for(int i=0;i<numOfRows;i++)
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = list.AddItem(itemCreateInfo);
                    foreach (var field in Fields)
                    {
                        if (field.FromBaseType == false && (!field.SchemaXml.Contains(" SourceID=\"http")))
                        {
                            Console.Write(field.InternalName);
                            if (_ints.Contains(field.FieldTypeKind))
                            {
                                newItem[field.InternalName] = 1;
                            }
                            else if (_bools.Contains(field.FieldTypeKind))
                            {
                                newItem[field.InternalName] = true;
                            }
                            else if (_strings.Contains(field.FieldTypeKind))
                            {
                                newItem[field.InternalName] = "Dummy";
                            }
                            else if (_datetime.Contains(field.FieldTypeKind))
                            {
                                newItem[field.InternalName] = DateTime.Now;
                            }
                            else if (_double.Contains(field.FieldTypeKind))
                            {
                                newItem[field.InternalName] = 1.5;
                            }
                        }
                    }
                    newItem["Title"] = "dummy";
                    newItem.Update();
                    context2.ExecuteQuery();
                }
                
            });

            /*
            foreach (var list in lists)
            {
            //var list = context.Web.Lists.GetByTitle("aditya");
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = list.AddItem(itemCreateInfo);
                var Fields = list.Fields;
                context.Load(Fields);
                context.ExecuteQuery();
                foreach(var field in Fields)
                {
                    if(field.FromBaseType == false && (!field.SchemaXml.Contains(" SourceID=\"http")))
                    {
                    Console.Write(field.InternalName);
                        if (_ints.Contains(field.FieldTypeKind))
                        {
                            newItem[field.InternalName] = 1;
                        }
                        else if(_bools.Contains(field.FieldTypeKind))
                        {
                            newItem[field.InternalName] = true;
                        }
                        else if (_strings.Contains(field.FieldTypeKind))
                        {
                            newItem[field.InternalName] = "Dummy";
                        }
                        else if (_datetime.Contains(field.FieldTypeKind))
                        {
                            newItem[field.InternalName] =DateTime.Now;
                        }
                        else if(_double.Contains(field.FieldTypeKind)) {
                            newItem[field.InternalName] = 1.5;
                        }
                    }
                }
            newItem["Title"] = "dummy";
                newItem.Update();
                context.ExecuteQuery();
            }

        }
            */
        }
        catch (Exception ex)
        {
            handleError(ex);
        }
        return true;
    }
}

