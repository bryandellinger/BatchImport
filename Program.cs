using ADBatchImportForInterProd;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;


namespace ADBatchImportforInterProd
{
    class Program
    {
        private readonly string connectionString;
        private readonly string adDomain;
        private readonly bool isFiltered;  // if is filtered is set to true in the config file then only employees in the filtered list will be added
        private readonly string[] filteredCompanies; // comma seperated list of companies or ExtensionAttribute2 located in config file
        private readonly string[] columnNames;
        private List<ADDepUser> recordList;
        private readonly IBatchService service;
       
        public Program(IBatchService batchservice)
        {
            service = batchservice;
            connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            adDomain = ConfigurationManager.AppSettings["AD_Domain"];
            isFiltered = ConfigurationManager.AppSettings["isFiltered"] == "true";
            filteredCompanies = ConfigurationManager.AppSettings["FilteredCompanies"].Split(',');
            columnNames = service.GetColumnNames();
            recordList = new List<ADDepUser>();
        }

        static void Main(string[] args)

        {
            var p = new Program(new BatchService());
            p.service.WriteToEventLog("Active Directory Extract -- Starting");
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, p.adDomain))
                {
                    using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                    {
                        Stopwatch stopWatch = new Stopwatch();
                        stopWatch.Start();
                        foreach (Principal result in searcher.FindAll())
                        {
                            if (p.recordList.Count % 10000 == 0 && p.recordList.Count > 0)
                            {
                                p.service.WriteToEventLog("Active Directory Extract -- Processed " + p.recordList.Count + " records in: " + p.service.getElapsedTime(stopWatch));
                            }

                            DirectoryEntry entry = result.GetUnderlyingObject() as DirectoryEntry;
                            if (
                                (entry.Properties["company"].Value != null || entry.Properties["extensionAttribute2"].Value != null) &&
                                entry.Properties["objectClass"].Value != null &&
                                !String.IsNullOrEmpty(entry.Properties["objectClass"].Value.ToString()) &&
                                ((IEnumerable)entry.Properties["objectClass"]).Cast<object>()
                                 .Select(x => x.ToString())
                                 .ToArray().Contains("user"))
                                    {
                                        p.recordList.Add(ADDepUser.GetADDepUserFromEntry(entry));
                                    }
                        }
                        p.service.WriteToEventLog("Active Directory Extract -- Processed " + p.recordList.Count + " records in: " + p.service.getElapsedTime(stopWatch));
                        stopWatch.Stop();
                    }
                }
                DataTable dt = new DataTable();

                foreach (string columnName in p.columnNames)
                    dt.Columns.Add(new DataColumn(columnName, typeof(string)));

                // if the is filtered is set to true in config file then only add the employees that belong to the companies in the list
                if (p.isFiltered)
                {
                    List<ADDepUser> list1 = (from r in p.recordList
                                             join c in p.filteredCompanies on r.Company equals c
                                             select r).ToList();


                    List<ADDepUser> list2 = (from r in p.recordList
                                             join c in p.filteredCompanies on r.ExtensionAttribute2 equals c
                                             select r).ToList();

                    p.recordList = list1;
                    p.recordList.AddRange(list2);
                }
                var distinctRecordList = p.recordList.GroupBy(x => x.ExtensionAttribute1).Select(x => x.First()).ToList(); // get rid of any duplicates

                foreach (var record in distinctRecordList)
                {
                    dt.Rows.Add(new string[]
                    {
                    p.service.GetDBValue(record.Cn),
                    p.service.GetDBValue(record.ExtensionAttribute1),
                    p.service.GetDBValue(record.ExtensionAttribute13),
                    p.service.GetDBValue(record.ExtensionAttribute3),
                    p.service.GetDBValue(record.ExtensionAttribute4),
                    p.service.GetDBValue(record.ExtensionAttribute7),
                    p.service.GetDBValue(record.GivenName),
                    p.service.GetDBValue(record.Initials),
                    p.service.GetDBValue(record.L),
                    p.service.GetDBValue(record.PostalCode),
                    p.service.GetDBValue(record.Sn),
                    p.service.GetDBValue(record.StreetAddress),
                    p.service.GetDBValue(record.TelephoneNumber),
                    p.service.GetDBValue(record.WhenChanged),
                    p.service.GetDBValue(record.WhenCreated),
                    });
                }


                p.service.ExecuteNonQuery("truncate table ad_dep_users");

                using (SqlBulkCopy bc = new SqlBulkCopy(p.connectionString))
                {
                    bc.DestinationTableName = "[dbo].[AD_DEP_USERS]";
                    bc.WriteToServer(dt);
                }

                p.service.ExecuteNonQuery("truncate table cwopa_agency_file");

                p.service.ExecuteNonQuery(p.service.getCWOPAInsertStatement());


                p.service.WriteToEventLog("Active Directory Extract -- Completed");
            }
            catch (Exception ex)
            {

                p.service.WriteToEventLog($"Active Directory Extract Error: {ex.Message}");
                throw;
            }

        }
    }
}
