using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureNotifications
{
    class SendAutomatedEmail
    {
        #region Env Variables
        string sValue = ConfigurationManager.AppSettings["TFSQuery1"];
        string smtpfrom = ConfigurationManager.AppSettings["smtpfrom"];
        string smtphost = ConfigurationManager.AppSettings["smtphost"];
        string smtpport = ConfigurationManager.AppSettings["smtpport"];
        string smtpuserName = ConfigurationManager.AppSettings["smtpuserName"];
        string smtppassword = ConfigurationManager.AppSettings["smtppassword"];
        string smtpto = ConfigurationManager.AppSettings["smtpto"];
        string TfsUrl = ConfigurationManager.AppSettings["TfsUrl"];
        string personalAccessToken = ConfigurationManager.AppSettings["personalAccessToken"];
        string emailSubject = ConfigurationManager.AppSettings["subject"];


        string col1 = ConfigurationManager.AppSettings["col1"];
        string col2 = ConfigurationManager.AppSettings["col2"];
        string col3 = ConfigurationManager.AppSettings["col3"];
        string col4 = ConfigurationManager.AppSettings["col4"];
        string col5 = ConfigurationManager.AppSettings["col5"];
        string col6 = ConfigurationManager.AppSettings["col6"];
        string col7 = ConfigurationManager.AppSettings["col7"];
        string col8 = ConfigurationManager.AppSettings["col8"];

        string iterationPath = String.Empty;
        #endregion

        public async Task<DataTable> CreateReportAsync(bool table1)
        {
            Uri orgUrl = new Uri(TfsUrl);
            VssConnection connection = new VssConnection(orgUrl, new VssBasicCredential(string.Empty, personalAccessToken));

            WorkItemTrackingHttpClient witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            Wiql wiql = new Wiql();

            wiql.Query = sValue;
            WorkItemQueryResult tasks = await witClient.QueryByWiqlAsync(wiql);

            IEnumerable<WorkItemReference> tasksRefs;
            tasksRefs = tasks.WorkItems.OrderBy(x => x.Id);

            List<WorkItem> tasksList = witClient.GetWorkItemsAsync(tasksRefs.Select(wir => wir.Id)).Result;
            List<String> taskDetails = new List<String>();

            #region TaskList to datatable
            DataTable table = new DataTable();

            table.Columns.Add(new DataColumn(col1, typeof(string)));
            table.Columns.Add(new DataColumn(col2, typeof(string)));
            table.Columns.Add(new DataColumn(col3, typeof(string)));
            table.Columns.Add(new DataColumn(col4, typeof(string)));
            table.Columns.Add(new DataColumn(col5, typeof(int)));
            table.Columns.Add(new DataColumn(col6, typeof(int)));
            table.Columns.Add(new DataColumn(col7, typeof(int)));
            table.Columns.Add(new DataColumn(col8, typeof(string)));


            foreach (var task in tasksList)
            {
                DataRow dataRow = null;

                dataRow = table.NewRow();
                if (task.Fields.ContainsKey("System." + col1))
                {
                    dataRow[col1] = (task.Fields["System." + col1].ToString());
                }
                else
                {
                    dataRow[col1] = String.Empty;
                }
                if (task.Fields.ContainsKey("System." + col2))
                {
                    dataRow[col2] = (task.Fields["System." + col2].ToString());
                }
                else
                {
                    dataRow[col2] = String.Empty;
                }
                if (task.Fields.ContainsKey("System." + col3))
                {
                    dataRow[col3] = (((IdentityRef)task.Fields["System." + col3]).DisplayName.ToString());
                }
                else
                {
                    dataRow[col3] = String.Empty;
                }
                if (task.Fields.ContainsKey("System." + col4))
                {
                    dataRow[col4] = (task.Fields["System." + col4].ToString());
                    iterationPath = (task.Fields["System." + col4].ToString());
                }
                else
                {
                    dataRow[col4] = String.Empty;
                }
                if (task.Fields.ContainsKey("Microsoft.VSTS.Scheduling." + col5))
                {
                    dataRow[col5] = (task.Fields["Microsoft.VSTS.Scheduling." + col5]);
                }
                else
                {
                    dataRow[col5] = 0;
                }
                if (task.Fields.ContainsKey("Microsoft.VSTS.Scheduling." + col6))
                {
                    dataRow[col6] = (task.Fields["Microsoft.VSTS.Scheduling." + col6]);
                }
                else
                {
                    dataRow[col6] = 0;
                }
                if (task.Fields.ContainsKey("Microsoft.VSTS.Scheduling." + col7))
                {
                    dataRow[col7] = (task.Fields["Microsoft.VSTS.Scheduling." + col7]);
                }
                else
                {
                    dataRow[col7] = 0;
                }
                if (task.Fields.ContainsKey("System." + col8))
                {
                    dataRow[col8] = (task.Fields["System." + col8].ToString());
                }
                else
                {
                    dataRow[col8] = String.Empty;
                }
                table.Rows.Add(dataRow);
            }
            #endregion
  
            var newDt = table.AsEnumerable()
            .GroupBy(r => r.Field<string>("AssignedTo"))
            .Select(g =>
            {
                var row = table.NewRow();

                row["AssignedTo"] = g.Key;
                row["WorkItemType"] = "Task";
                  // row["IterationPath"] = 
                  row["OriginalEstimate"] = g.Sum(r => r.Field<int>("OriginalEstimate"));
                row["CompletedWork"] = g.Sum(r => r.Field<int>("CompletedWork"));
                row["RemainingWork"] = g.Sum(r => r.Field<int>("RemainingWork"));

                return row;
            }).CopyToDataTable();


            newDt.Columns.Remove("Title");
            newDt.Columns.Remove("IterationPath");
            newDt.Columns.Remove("State");


            DataTable table2 = new DataTable();

            table2.Columns.Add(new DataColumn(col4, typeof(string)));
            table2.Columns.Add(new DataColumn(col5, typeof(int)));
            table2.Columns.Add(new DataColumn(col6, typeof(int)));
            table2.Columns.Add(new DataColumn(col7, typeof(int)));


            DataRow dataRow2 = null;
            dataRow2 = table2.NewRow();
            dataRow2[col4] = iterationPath.ToString();
            dataRow2[col5] = Convert.ToInt32(table.Compute("SUM(OriginalEstimate)", string.Empty));
            dataRow2[col6] = Convert.ToInt32(table.Compute("SUM(CompletedWork)", string.Empty));
            dataRow2[col7] = Convert.ToInt32(table.Compute("SUM(RemainingWork)", string.Empty));
            table2.Rows.Add(dataRow2);

            if (table1)
            {
                return table;
            }
            else
            {
                return table2;
            }
            }
    }

}


