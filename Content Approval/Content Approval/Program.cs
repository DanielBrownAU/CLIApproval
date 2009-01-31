using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace DanielBrown.SharePoint.Tools.CLI
{
    class Program
    {
        private static void DisplayUsage()
        {
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Options:");
            Console.WriteLine("-enableversioning\tEnable versioning on all document libraries");
            Console.WriteLine("-disableversioning\tdisable versioning on all document libraries");
            Console.WriteLine("-enablecontentapproval\tEnable Content Approval on all document libraries");
            Console.WriteLine("-disablecontentapproval\tDisable Content Approval on all document libraries");
            Console.WriteLine("-includesubsites\tApplies settings to sub sites");
        }
        static void Main(string[] args)
        {
            string requestUrl = string.Empty;

            bool ContentApproval = false;
            bool Versioning = false;
            bool IsRecursive = false;

            Console.WriteLine("Mass enabling or disabling of Content Approval and Versioning on a sites document libraries.");
            Console.WriteLine("Developed by Daniel Brown (daniel.brown@internode.on.net)");

            // Check if we have command line args
            if ((args == null) || (args.Length == 0))
            {
                // If we do not, display the usgae
                DisplayUsage();
                return;
            }
            else
            {
                // If we do....

                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("Using Options:");
                // Loop though each arg and set out settings
                foreach (string arg in args)
                {
                    // Split any Args which use = (i.e. url=http://localhost )
                    string[] ArgCollection = arg.Split('=');

                    // Key is the first part of the Key/Pair
                    string Key = ArgCollection[0].ToLower();

                    // Value is the second part of the Key/Pair
                    string Value = string.Empty;

                    // Check we have the right amount of args in the split and that its not null or empty
                    if ((ArgCollection.Length == 2) && (!string.IsNullOrEmpty(ArgCollection[1])))
                    {
                        Value = ArgCollection[1].ToLower();
                    }

                    switch (Key)
                    {
                        case "-url":
                            if (string.IsNullOrEmpty(Value))
                            {
                                Console.WriteLine("You need ot enter a URL");
                                return;
                            }
                            else
                            {
                                requestUrl = Value;
                                Console.WriteLine("- Using site : {0}", requestUrl);
                            }
                            break;

                        case "-enableversioning": // Versioning: Enable
                            Versioning = true;
                            Console.WriteLine("- Enabling versioning");
                            break;

                        case "-disableversioning": // Versioning: Disable
                            Versioning = false;
                            Console.WriteLine("- Disabling Versioning");
                            break;

                        case "-enablecontentapproval": // Content Approval: Enable
                            ContentApproval = true;
                            Console.WriteLine("- Enabling Content Approval");
                            break;

                        case "-disablecontentapproval": // Content Approval: Disable
                            ContentApproval = false;
                            Console.WriteLine("- Disabling Content Approval");
                            break;

                        case "-includesubsites": // Recursive: Enable
                            IsRecursive = false;
                            Console.WriteLine("- Including Subsites");
                            break;
                        default:
                            Console.WriteLine("Unknown arg");
                            return;
                            break;
                    }
                }
            }

            // Everythign shoudl be valid at this point
            Process(requestUrl, ContentApproval, Versioning, IsRecursive);
            Console.WriteLine("Completed.");

        }

        private static void Process(string requestUrl, bool ContentApproval, bool Versioning, bool IsRecursive)
        {
            SPSite rootsite = null;
            SPWeb rootweb = null;

            try
            {
                // Open Site
                using (rootsite = new SPSite(requestUrl))
                {
                    using (rootweb = rootsite.OpenWeb())
                    {
                        // Allow unsafe updates
                        rootweb.AllowUnsafeUpdates = true;
                        rootweb.Update();

                        // Process settings of the document libraries
                        HandleDocumentLibrary(rootweb, ContentApproval, Versioning);

                        // disable unsafe updates
                        rootweb.AllowUnsafeUpdates = false;
                        rootweb.Update();

                        if (IsRecursive) // Check if recurve was defined
                        {
                            // Loop though each sub-site and handle versioning & content approval settings
                            for(int i = 0; i <= rootweb.Webs.Count; i++)
                            {
                                using (SPWeb subweb = rootweb.Webs[i])
                                {
                                    HandleDocumentLibrary(subweb, ContentApproval, Versioning);
                                }
                            }
                        }

                    }
                }
            }
            catch (SPException spex)
            {
                Console.WriteLine(spex.ToString());
                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("Push any key to continue");
                Console.ReadLine();
            }
            finally
            {
                // If somethign went south and rootweb is not null, dispose of it correctly
                if (rootweb != null)
                {
                    rootweb.Close();
                    rootweb.Dispose();
                }

                // If something went south and rootsite is not null, dispose of it correctly
                if (rootsite != null)
                {
                    rootsite.Close();
                    rootsite.Dispose();
                }
            }
        }

        private static void HandleDocumentLibrary(SPWeb web, bool ContentApproval, bool Versioning)
        {
            foreach (SPList list in web.Lists)
            {
                if (list.BaseTemplate == SPListTemplateType.DocumentLibrary)
                {
                    SPDocumentLibrary doclib = list as SPDocumentLibrary;

                    // Set "Content Approval"
                    doclib.EnableModeration = ContentApproval;

                    // Set "Versioning"
                    doclib.EnableVersioning = Versioning;

                    doclib.Update();
                }
            }
        }
    }
}