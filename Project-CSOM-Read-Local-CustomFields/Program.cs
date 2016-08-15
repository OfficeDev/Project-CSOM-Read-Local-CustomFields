using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;

/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
namespace ProjectReadLocalCustomFields
{
    class CSOM_LCF
    {
        // Project name to access
        private static String SampleProjectName = "Local Custom Fields";

        private static ProjectContext projContext =
            new ProjectContext("https://contoso.sharepoint.com/sites/pwa");

        static void Main(string[] args)
        {
            // Lists out the LCFs for a specific project in the PWA instance, consisting of the 
            // field type, field name, whether a lookup table is used, and the value and description of each each LCF entry. 

            // This app does the following:
            // 1. Retrieves a specific project, task(s), and custom fields associated with the tasks in the project
            // 2. Place the custom field key/value pairs in a dictionary
            // 3. Do the following for each Local Custom Field:
            //    A. Distinguish between simple values and values that use a lookup table for 
            //       the friendly values.
            //    B. Filter out partial entries.
            //    C. List the friendly value for the user (simple values)
            //    D. List the friendly value for the user (lookup tables)

            using (projContext)
            {
                // Supply user credentials

                SecureString passWord = new SecureString();
              foreach (char c in "password".ToCharArray()) passWord.AppendChar(c);


              projContext.Credentials = new SharePointOnlineCredentials("sarad@contoso.onmicrosoft.com", passWord);

                // 1. Retrieve the project, tasks, etc.
                var projColl = projContext.LoadQuery(projContext.Projects
                    .Where(p => p.Name == SampleProjectName)
                    .Include(
                        p => p.Id,
                        p => p.Name,
                        p => p.Tasks,
                        p => p.Tasks.Include(
                            t => t.Id,
                            t => t.Name,
                            t => t.CustomFields,
                            t => t.CustomFields.IncludeWithDefaultProperties(
                                cf => cf.LookupTable,
                                cf => cf.LookupEntries
                            )
                        )
                    )
                );

                projContext.ExecuteQuery();

                PublishedProject theProj = projColl.First();


                Console.WriteLine("Name:\t{0}", theProj.Name);
                Console.WriteLine("Id:\t{0}", theProj.Id);
                Console.WriteLine("Tasks count: {0}", theProj.Tasks.Count);
                Console.WriteLine("  -----------------------------------------------------------------------------");


                PublishedTaskCollection taskColl = theProj.Tasks;

                PublishedTask theTask = taskColl.First();

                CustomFieldCollection LCFColl = theTask.CustomFields;

                // 2. Place the task-specific custom field key/value pairs in a dictionary
                Dictionary<string, object> taskCF_Dict = theTask.FieldValues;

                if (LCFColl.Count > 0)
                {

                    Console.WriteLine("\n\tType\t   Name\t\t  L.UP   Value                  Description");
                    Console.WriteLine("\t--------   ------------   ----   --------------------   -----------");

                    // 3. For each custom field, do the follwoing:
                    foreach (CustomField cf in LCFColl)
                    {
                        // 3A. Distinguish LCF values that are simple from those that use lookup tables.
                        if (!cf.LookupTable.ServerObjectIsNull.HasValue ||
                                            (cf.LookupTable.ServerObjectIsNull.HasValue && cf.LookupTable.ServerObjectIsNull.Value))
                        {
                            if (taskCF_Dict[cf.InternalName] == null)
                            {   // 3B. Partial implementation. Not usable.
                                String textValue = "is not set";
                                Console.WriteLine("\t{0} {1}  {2}", cf.FieldType, cf.Name, textValue);
                            }
                            else     // 3C. Friendly value for the user (simple).
                            {
                                String textValue = taskCF_Dict[cf.InternalName].ToString();
                                Console.WriteLine("\t{0, -8}   {1, -15}       {2}", cf.FieldType, cf.Name, textValue);
                            }
                        }
                        else         //3D. Friendly value for the user that uses a Lookup table.
                        {

                            Console.Write("\t{0, -8}   {1, -15}", cf.FieldType, cf.Name);

                            String[] entries = (String[])taskCF_Dict[cf.InternalName];

                            foreach (String entry in entries)
                            {
                                var luEntry = projContext.LoadQuery(cf.LookupTable.Entries
                                        .Where(e => e.InternalName == entry));

                                projContext.ExecuteQuery();

                                Console.WriteLine("Yes    {0, -22}  {1}", luEntry.First().FullValue, luEntry.First().Description);
                            }
                        }


                    }   // End foreach CustomField

                }

            }

            Console.Write("\nPress any key to exit: ");
            Console.ReadKey(false);

        }   // End of Main

    }
}

