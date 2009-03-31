using System;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;


namespace DanielBrown.SharePoint.Tools.CLI.SPCMA.STSADM
{
    public class SPMCACommand : ISPStsadmCommand
    {
        /// <summary>
        /// Shows the Help Message of the stsadm -o
        /// </summary>
        /// <param name="command">The command the user is getting help on</param>
        /// <returns>The Help message</returns>
        public string GetHelpMessage(string command)
        {
            command = command.ToLowerInvariant();

            StringBuilder sbHelp = new StringBuilder();

            sbHelp.AppendLine(Environment.NewLine);
            sbHelp.AppendLine("Mass enabling or disabling of Content Approval and Versioning on a sites document libraries.");
            sbHelp.AppendLine("Developed by Daniel Brown (daniel.brown@internode.on.net)");
            sbHelp.AppendLine(Environment.NewLine);

            switch (command)
            {         
                case "enableversioning":
                    {
                        sbHelp.AppendLine("-enableversioning\tEnable versioning on all document libraries");
                        sbHelp.AppendLine("Exmaple: stsadm -o enableversioning -url http://localhost -includesubsites");
                        break;
                    }
                case "disableversioning":
                    {
                        sbHelp.AppendLine("-disableversioning\tdisable versioning on all document libraries");
                        sbHelp.AppendLine("Exmaple: stsadm -o disableversioning -url http://localhost -includesubsites");
                        break;
                    }
                case "enablecontentapproval":
                    {
                        sbHelp.AppendLine("-enablecontentapproval\tEnable Content Approval on all document libraries");
                        sbHelp.AppendLine("Exmaple: stsadm -o enablecontentapproval -url http://localhost -includesubsites");
                        break;
                    }
                case "disablecontentapproval":
                    {
                        sbHelp.AppendLine("Exmaple: stsadm -o enablecontentapproval -url http://localhost -includesubsites");
                        sbHelp.AppendLine("-disablecontentapproval\tDisable Content Approval on all document libraries");
                        break;
                    }
                case "spmcahelp":
                    {
                        sbHelp.AppendLine("Options:");
                        sbHelp.AppendLine("stsadm -help enableversioning\tEnable versioning on all document libraries");
                        sbHelp.AppendLine("stsadm -help disableversioning\tdisable versioning on all document libraries");
                        sbHelp.AppendLine("stsadm -help enablecontentapproval\tEnable Content Approval on all document libraries");
                        sbHelp.AppendLine("stsadm -help disablecontentapproval\tDisable Content Approval on all document libraries");
                    }
                    break;
                default:
                    {
                        throw new InvalidOperationException();
                    }
            }

            sbHelp.AppendLine(Environment.NewLine);
            sbHelp.AppendLine("Other:");
            sbHelp.AppendLine("-url <full url to a site in SharePoint>");
            sbHelp.AppendLine("-includesubsites will loop though al subsites from the url supplied.");
            sbHelp.AppendLine(Environment.NewLine);
            sbHelp.AppendLine("NOTE: -url is required");

            return sbHelp.ToString();
        }

        /// <summary>
        /// Runs the command entered by the user
        /// </summary>
        /// <param name="command">The name of the custom operation.</param>
        /// <param name="keyValues">The parameters, if any, that are added to the command line.s</param>
        /// <param name="output">An output string, if needed.</param>
        /// <returns>An Int32 that can be used to signal the result of the operation.
        /// For proper operation with STSADM, use the following rules when implementing.
        /// Return 0 for success. 
        /// Return GeneralError (-1) for any error other than syntax. 
        /// Return SyntaxError (-2) for a syntax error. 
        /// When 0 is returned, STSADM streams output, if it is not null, to the console. 
        /// When SyntaxError is returned, STSADM calls GetHelpMessage and its return value is streamed to the console, and it streams output, 
        /// if it is not null, to stderr (standard error). 
        /// To obtain the content of stderr in managed code, use Error. When any other value is returned, STSADM streams output, 
        /// if it is not null, to stderr.</returns>
        public int Run(string command, StringDictionary keyValues, out string output)
        {
            command = command.ToLowerInvariant();

            bool ContentApproval = false;
            bool Versioning = false;

            switch (command)
            {
                case "enableversioning":
                    {
                        Versioning = true;
                        break;
                    }
                case "disableversioning":
                    {
                        Versioning = false;
                        break;
                    }
                case "enablecontentapproval":
                    {
                        ContentApproval = true;
                        break;
                    }
                case "disablecontentapproval":
                    {
                        ContentApproval = false;
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException();
                    }
            }

            // Pass over to the Process method to actualy process our command
            return this.Process(keyValues, out output, ContentApproval, Versioning);
        }

        /// <summary>
        /// Checks for a valid -url
        /// </summary>
        /// <param name="keyValues">The parameters, if any, that are added to the command line.s</param>
        /// <returns>the value of the -url token</returns>
        private string GetURL(StringDictionary keyValues)
        {
            // check for URL
            if (!keyValues.ContainsKey("url"))
            {
                throw new InvalidOperationException("The url parameter was not specified.");
            }

            // check if the value is valid
            if (string.IsNullOrEmpty(keyValues["url"]))
            {
                throw new InvalidOperationException("The url parameter was invalid..");
            }

            // return the value
            return keyValues["url"];
        }

        private int Process(StringDictionary keyValues, out string output, bool ContentApproval, bool Versioning)
        {
            // Get the URL
            string requestUrl = this.GetURL(keyValues);

            bool IsRecursive = false;

            // Set if we are running Recursive or not
            if (string.IsNullOrEmpty(keyValues["includesubsites"]))
            {
                IsRecursive = true;
            }

            SPSite rootsite = default(SPSite);
            SPWeb rootweb = null;

            try
            {
                Uri rootweburi = new Uri(requestUrl);

                if (SPSite.Exists(rootweburi)) // Check if Site is there
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
                            this.HandleDocumentLibrary(rootweb, ContentApproval, Versioning);

                            // disable unsafe updates
                            rootweb.AllowUnsafeUpdates = false;
                            rootweb.Update();

                            if (IsRecursive) // Check if recurve was defined
                            {
                                // Loop though each sub-site and handle versioning & content approval settings
                                for (int i = 0; i <= rootweb.Webs.Count; i++)
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
                else
                {

                    Console.WriteLine(string.Format("Web Application at {0} was not found, please check the \"URL\" argument and try again.", requestUrl));
                }
            }
            catch (SPException spex)
            {
                StringBuilder sbError = new StringBuilder();
                sbError.AppendLine(spex.ToString());
                output = sbError.ToString();
                return -1;
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

            output = "Completed.";
            return 0;
        }

        private void HandleDocumentLibrary(SPWeb web, bool ContentApproval, bool Versioning)
        {
            foreach (SPList list in web.Lists)
            {
                if (list.BaseTemplate == SPListTemplateType.DocumentLibrary)
                {
                    Console.WriteLine("Processing " + list.Title);

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