using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;

namespace AzureNotifications
{
     static class Program
    {
        static async Task Main(string[] args)
        {
            try
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

                #region WorkEstimates Table
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

                //prepare chart control...
                CreateChart(newDt,true);
                #endregion

                #region State Table
                var newDt2 = table.AsEnumerable()
              .GroupBy(r => r.Field<string>("State"))
              .Select(g =>
              {
                  var row = table.NewRow();

                  row["State"] = g.Key;
                  // row["IterationPath"] = 
                  row["OriginalEstimate"] = g.Sum(r => r.Field<int>("OriginalEstimate"));
                  row["CompletedWork"] = g.Sum(r => r.Field<int>("CompletedWork"));
                  row["RemainingWork"] = g.Sum(r => r.Field<int>("RemainingWork"));

                  return row;
              }).CopyToDataTable();

                newDt2.Columns.Remove("Title");
                newDt2.Columns.Remove("WorkItemType");
                newDt2.Columns.Remove("IterationPath");
                newDt2.Columns.Remove("AssignedTo");

                CreateChart(newDt2,false);
                #endregion


                if (table.Rows.Count > 0)
                {
                    #region SendEmail
                    SendEmail(smtpfrom, smtphost, smtpport, smtpuserName, smtppassword, smtpto, newDt, table2, emailSubject);
                    #endregion
                }
            }
            catch (Exception Ex)
            {
                string Current_directory_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));
                if (!File.Exists(Current_directory_path))
                {
                    File.Create(Current_directory_path + "\\log.txt").Dispose();
                }
                using (StreamWriter sw = File.AppendText(Current_directory_path + "\\log.txt"))
                {
                    string error = "Log Written Date:" + " " + DateTime.Now.ToString() + "Error Message:" + " " + Ex.InnerException + "Exception:" + " " + Ex.Message;
                    sw.WriteLine(error);
                    sw.WriteLine("|||");
                    sw.Flush();
                    sw.Close();
                }
            }

        }

        private static void SendEmail(string smtpfrom, string smtphost, string smtpport, string smtpuserName, string smtppassword, string smtpto, DataTable table, DataTable table2, string emailSubject)
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient(smtphost);

            mail.From = new MailAddress(smtpfrom);
            mail.To.Add(smtpto);
            mail.Subject = emailSubject;
            mail.Body = "<br>";
            mail.Body += ConvertDataTableToHTML(table2);
            mail.Body += "<br>";
            mail.Body += ConvertDataTableToHTML(table);
            mail.IsBodyHtml = true;

            LinkedResource LinkedImage1 = new LinkedResource(@"c:\myChart1.png");
            LinkedImage1.ContentId = "MyPic1";
            LinkedImage1.ContentType = new ContentType(MediaTypeNames.Image.Jpeg);

            LinkedResource LinkedImage2 = new LinkedResource(@"c:\myChart2.png");
            LinkedImage2.ContentId = "MyPic2";
            LinkedImage2.ContentType = new ContentType(MediaTypeNames.Image.Jpeg);

            LinkedResource LinkedImage3 = new LinkedResource(@"c:\myChart3.png");
            LinkedImage3.ContentId = "MyPic3";
            LinkedImage3.ContentType = new ContentType(MediaTypeNames.Image.Jpeg);


            AlternateView htmlView1 = AlternateView.CreateAlternateViewFromString(
             "<center><b> Original Estimate<center> <b><img src=cid:MyPic1> <br> <center> Completed Work<center> <img src=cid:MyPic2> <br> <center>Remaining Work <center> <img src=cid:MyPic3> " + mail.Body, null, MediaTypeNames.Text.Html);

            //AlternateView htmlView2 = AlternateView.CreateAlternateViewFromString(
            // "Completed Work. <img src=cid:MyPic2>", null, MediaTypeNames.Text.Html);

            //AlternateView htmlView3 = AlternateView.CreateAlternateViewFromString(
            // "Remaining Work. <img src=cid:MyPic3>", null, MediaTypeNames.Text.Html);

            htmlView1.LinkedResources.Add(LinkedImage1);
            htmlView1.LinkedResources.Add(LinkedImage2);
            htmlView1.LinkedResources.Add(LinkedImage3);
            mail.AlternateViews.Add(htmlView1);


            SmtpServer.Port = Convert.ToInt32(smtpport);
            SmtpServer.Credentials = new System.Net.NetworkCredential(smtpuserName, smtppassword);
            //SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
        }

        private static void CreateChart(DataTable table , bool isWorkItem)
        {
            if (isWorkItem)
            {
            Chart chart = new Chart();
            chart.DataSource = table;
            chart.Width = 600;
            chart.Height = 350;
            //create serie...
            Series serie1 = new Series();
            serie1.Name = "Serie1";
            serie1.Color = Color.FromArgb(112, 255, 200);
            serie1.BorderColor = Color.FromArgb(164, 164, 164);
            serie1.ChartType = SeriesChartType.Column;
            serie1.BorderDashStyle = ChartDashStyle.Solid;
            serie1.BorderWidth = 1;
            serie1.ShadowColor = Color.FromArgb(128, 128, 128);
            serie1.ShadowOffset = 1;
            serie1.IsValueShownAsLabel = true;
            serie1.XValueMember = "AssignedTo";
            serie1.YValueMembers = "OriginalEstimate";
            serie1.Font = new Font("Tahoma", 8.0f);
            serie1.BackSecondaryColor = Color.FromArgb(0, 102, 153);
            serie1.LabelForeColor = Color.FromArgb(100, 100, 100);
            chart.Series.Add(serie1);
            //create chartareas...
            ChartArea ca = new ChartArea();
            ca.Name = "ChartArea1";
            ca.BackColor = Color.White;
            ca.BorderColor = Color.FromArgb(26, 59, 105);
            ca.BorderWidth = 0;
            ca.BorderDashStyle = ChartDashStyle.Solid;
            ca.AxisX = new Axis();
            ca.AxisY = new Axis();
            ca.AxisX.Interval = 1;
            ca.AxisX.MajorGrid.LineWidth = 0;
            ca.AxisY.MajorGrid.LineWidth = 0;
            chart.ChartAreas.Add(ca);
            //databind...
            chart.DataBind();
            //save result...
            chart.SaveImage(@"c:\myChart1.png", ChartImageFormat.Png);

            Chart chart2 = new Chart();
            chart2.DataSource = table;
            chart2.Width = 600;
            chart2.Height = 350;
            //create serie...
            Series serie2 = new Series();
            serie2.Name = "Serie1";
            serie2.Color = Color.FromArgb(112, 255, 200);
            serie2.BorderColor = Color.FromArgb(164, 164, 164);
            serie2.ChartType = SeriesChartType.Column;
            serie2.BorderDashStyle = ChartDashStyle.Solid;
            serie2.BorderWidth = 1;
            serie2.ShadowColor = Color.FromArgb(128, 128, 128);
            serie2.ShadowOffset = 1;
            serie2.IsValueShownAsLabel = true;
            serie2.XValueMember = "AssignedTo";
            serie2.YValueMembers = "CompletedWork";
            serie2.Font = new Font("Tahoma", 8.0f);
            serie2.BackSecondaryColor = Color.FromArgb(0, 102, 153);
            serie2.LabelForeColor = Color.FromArgb(100, 100, 100);
            chart2.Series.Add(serie2);
            //create chartareas...
            ChartArea ca2 = new ChartArea();
            ca2.Name = "ChartArea1";
            ca2.BackColor = Color.White;
            ca2.BorderColor = Color.FromArgb(26, 59, 105);
            ca2.BorderWidth = 0;
            ca2.BorderDashStyle = ChartDashStyle.Solid;
            ca2.AxisX = new Axis();
            ca2.AxisY = new Axis();
            ca2.AxisX.Interval = 1;
            ca2.AxisX.MajorGrid.LineWidth = 0;
            ca2.AxisY.MajorGrid.LineWidth = 0;
            chart2.ChartAreas.Add(ca2);
            //databind...
            chart2.DataBind();
            //save result...
            chart2.SaveImage(@"c:\myChart2.png", ChartImageFormat.Png);

            Chart chart3 = new Chart();
            chart3.DataSource = table;
            chart3.Width = 600;
            chart3.Height = 350;
            //create serie...
            Series serie3 = new Series();
            serie3.Name = "Serie1";
            serie3.Color = Color.FromArgb(112, 255, 200);
            serie3.BorderColor = Color.FromArgb(164, 164, 164);
            serie3.ChartType = SeriesChartType.Column;
            serie3.BorderDashStyle = ChartDashStyle.Solid;
            serie3.BorderWidth = 1;
            serie3.ShadowColor = Color.FromArgb(128, 128, 128);
            serie3.ShadowOffset = 1;
            serie3.IsValueShownAsLabel = true;
            serie3.XValueMember = "AssignedTo";
            serie3.YValueMembers = "RemainingWork";
            serie3.Font = new Font("Tahoma", 8.0f);
            serie3.BackSecondaryColor = Color.FromArgb(0, 102, 153);
            serie3.LabelForeColor = Color.FromArgb(100, 100, 100);
            chart3.Series.Add(serie3);
            //create chartareas...
            ChartArea ca3 = new ChartArea();
            ca3.Name = "ChartArea1";
            ca3.BackColor = Color.White;
            ca3.BorderColor = Color.FromArgb(26, 59, 105);
            ca3.BorderWidth = 0;
            ca3.BorderDashStyle = ChartDashStyle.Solid;
            ca3.AxisX = new Axis();
            ca3.AxisY = new Axis();
            ca3.AxisX.Interval = 1;
            ca3.AxisX.MajorGrid.LineWidth = 0;
            ca3.AxisY.MajorGrid.LineWidth = 0;
            chart3.ChartAreas.Add(ca3);
            //databind...
            chart3.DataBind();
            //save result...
            chart3.SaveImage(@"c:\myChart3.png", ChartImageFormat.Png);
            }
            else
            {
                Chart chart3 = new Chart();
                chart3.DataSource = table;
                chart3.Width = 600;
                chart3.Height = 350;
                //create serie...
                Series serie3 = new Series();
                serie3.Name = "Serie1";
                serie3.Color = Color.FromArgb(112, 255, 200);
                serie3.BorderColor = Color.FromArgb(164, 164, 164);
                serie3.ChartType = SeriesChartType.Column;
                serie3.BorderDashStyle = ChartDashStyle.Solid;
                serie3.BorderWidth = 1;
                serie3.ShadowColor = Color.FromArgb(128, 128, 128);
                serie3.ShadowOffset = 1;
                serie3.IsValueShownAsLabel = true;
                serie3.XValueMember = "State";
                serie3.YValueMembers = "RemainingWork";
                serie3.Font = new Font("Tahoma", 8.0f);
                serie3.BackSecondaryColor = Color.FromArgb(0, 102, 153);
                serie3.LabelForeColor = Color.FromArgb(100, 100, 100);
                chart3.Series.Add(serie3);
                //create chartareas...
                ChartArea ca3 = new ChartArea();
                ca3.Name = "ChartArea1";
                ca3.BackColor = Color.White;
                ca3.BorderColor = Color.FromArgb(26, 59, 105);
                ca3.BorderWidth = 0;
                ca3.BorderDashStyle = ChartDashStyle.Solid;
                ca3.AxisX = new Axis();
                ca3.AxisY = new Axis();
                ca3.AxisX.Interval = 1;
                ca3.AxisX.MajorGrid.LineWidth = 0;
                ca3.AxisY.MajorGrid.LineWidth = 0;
                chart3.ChartAreas.Add(ca3);
                //databind...
                chart3.DataBind();
                //save result...
                chart3.SaveImage(@"c:\myChart4.png", ChartImageFormat.Png);
            }
        }

        public static string ConvertDataTableToHTML(DataTable dt)
        {
            string html = "<table border = '1'>";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<td>" + dt.Columns[i].ColumnName + "</td>";
            html += "</tr>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }


    }
}
