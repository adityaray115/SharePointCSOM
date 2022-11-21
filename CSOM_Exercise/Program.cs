using Microsoft.IdentityModel.Tokens;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.PortalLaunch;
using System;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using static System.Net.Mime.MediaTypeNames;

namespace CSOMDemo
{
    class Program
    {
        public static void Main(String[] Args)
        {
            if (CSOMHandler.connetToSharePoint() == false)
            {
                exit();
                return;
            }
            
            //CSOMHandler.dummy_func();
            //return;

            printMenu();
            string choice=Console.ReadLine();

            if (choice.IsNullOrEmpty())
                choice = "Invalid";



            while(true)
            {
                bool success = true;
                switch(choice)
                {
                    case "1":
                        
                        success=showAllListsAndLibraries();
                        if (success == false)
                        {
                            exit();
                            return;
                        }
                        break;
                    case "2":
                        success = listsAndLibrariesWithMoreThanNItems();
                        if (success == false)
                        {
                            exit();
                            return;
                        }
                        break;
                    case "3":
                        success = addItemsToListAndLibraries();
                        if (success == false)
                        {
                            exit();
                            return;
                        }
                        break;
                    case "4":
                        success = selectiveSearch();
                        if (success == false)
                        {
                            exit();
                            return;
                        }
                        break;
                    case "5":
                        Console.WriteLine("Exiting Application");
                        return;
                    default:
                        Console.WriteLine("Invalid Input");
                        break;
                }

                if(!success)
                {
                    exit();
                    return;
                }

                printMenu();
                choice= Console.ReadLine();
                if (choice.IsNullOrEmpty())
                    choice = "Invalid";
            }
        }

        private static bool selectiveSearch()
        {
            try
            {
                return CSOMHandler.SelectiveSearch();
            }
            catch(Exception ex)
            {
                CSOMHandler.handleError(ex);
                return false;
            }
        }

        private static bool addItemsToListAndLibraries()
        {
            try
            {
                return CSOMHandler.addMItems();
            }
            catch( Exception ex)
            {
                CSOMHandler.handleError(ex);
                return false;
            }
            
        }

        private static bool listsAndLibrariesWithMoreThanNItems()
        {
            try
            {
                Console.Write("Enter Minimum entry count for the list:  ");
                string input = Console.ReadLine();
                if(!int.TryParse(input,out int N))
                {
                    N = 0;
                }
                if(N<0)
                {
                    Console.WriteLine("Negative Value Entered");
                    return true;
                }
                return CSOMHandler.listWithMoreThanN(N);
            }
            catch (Exception ex)
            {
                CSOMHandler.handleError(ex);
                return false;
            }
            return true;
        }

        private static bool showAllListsAndLibraries()
        {
            try
            {
                if (CSOMHandler.showAllListsAndLibraries() == false)
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                CSOMHandler.handleError(e);
                return false;
            }
            return true;
        }

        private static void exit()
        {
            Console.WriteLine("Error Occured.");
            Console.WriteLine("Exiting Application");
        }

        private static void printMenu()
        {
            Console.WriteLine("Select from the Menu:");
            Console.WriteLine("1: Show all Lists/Libraries");
            Console.WriteLine("2: List/Libraries with N or more Items");
            Console.WriteLine("3: Add item to all Lists/Libraries");
            Console.WriteLine("4: Selective Columns Search");
            Console.WriteLine("5: Exit");
        }
    }
}