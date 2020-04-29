using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ClosedXML.Excel;
using System.IO;
using System.Data.SqlClient;
using System.Drawing;

namespace SharePointOperations
{
    public partial class Default : System.Web.UI.Page
    {
        public static DataTable dt = null;
        public static string tableStructure = null;
        public static int count = 0;

        protected void Page_Load(object sender, EventArgs e)
        {




        }

        protected void btnDisplayGroups_Click(object sender, EventArgs e)
        {
            if (txtUrl.Text != "" || txtUrl.Text != string.Empty)
            {

                List<string> lst = LoadSharepointListItems(txtUrl.Text);

                chksharePointGroups.DataSource = lst;
                chksharePointGroups.DataBind();
                rwchkMessage.Visible = true;
                rwselformatte.Visible = true;
                rwExportToDb.Visible = true;


            }
            else
            {

                lblRes.Text = "Please enter the URL";
                rwchkMessage.Visible = false;
                rwselformatte.Visible = false;

                lblRes.ForeColor = System.Drawing.Color.Red;
            }
        }

        public List<string> LoadSharepointListItems(string siteUrl, bool isSharePOintOnline = false)
        {

            string gpNames = string.Empty;
            List<string> lstGroups = new List<string>();
            try
            {

                ClientContext clientContext1 = new ClientContext(siteUrl);
                Web oWebsite = clientContext1.Web;
                ListCollection collList = oWebsite.Lists;

                clientContext1.Load(collList);

                clientContext1.ExecuteQuery();

                foreach (List oList1 in collList)
                {
                    lstGroups.Add(oList1.Title);

                }

                return lstGroups;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("404"))
                {
                    lblRes.Text = "Please enter the correct URL or check the user permission on the site";
                }
                else
                {
                    lblRes.Text = ex.Message;

                }
                rwchkMessage.Visible = false;
                rwselformatte.Visible = false;
                lblRes.ForeColor = Color.Red;
                return lstGroups;
            }
        }
        protected void ddlSelectAuth_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (ddlSelectAuth.SelectedValue == "Windows")
            {
                trUserName.Visible = false;
                trpassword.Visible = false;

            }
            else
            {
                trUserName.Visible = true;
                trpassword.Visible = true;

            }
        }
        protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            string selValue = DropDownList1.SelectedValue;
            if (selValue.ToUpper() == "SQLDATABASE")
            {
                rwDataBase.Visible = true;
                rwDbdatasource.Visible = true;
                rwExportToDb.Visible = true;
                rwAuthentication.Visible = true;

            }
            else
            {

                rwDataBase.Visible = false;
                rwDbdatasource.Visible = false;
                rwExportToDb.Visible = true;
                rwAuthentication.Visible = false;
                trUserName.Visible = false;
                trpassword.Visible = false;
                
                
            }
        }

        public void ReadData(string listname, string selValue)
        {
            try
            {
                ClientContext clientContext = new ClientContext(txtUrl.Text);
                List oList = clientContext.Web.Lists.GetByTitle(listname);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "";
                int i = 0;
                int j = 0;
                Microsoft.SharePoint.Client.ListItemCollection collListItem = oList.GetItems(camlQuery);
                clientContext.Load(collListItem);
                clientContext.ExecuteQuery();
                string colName = string.Empty;

                foreach (Microsoft.SharePoint.Client.ListItem oListItem in collListItem)
                {
                    count = oListItem.FieldValues.Count;
                    foreach (var f in oListItem.FieldValues)
                    {
                        colName = colName + f.Key + ",";
                        i = i + 1;
                        if (i == count)
                        {
                            break;
                        }
                    }

                    break;
                }
                colName = colName.TrimEnd(',');
                string[] colnames = colName.Split(',');
                createTable(colnames, selValue, listname);
                string key = string.Empty;
                string value = string.Empty;
                foreach (Microsoft.SharePoint.Client.ListItem oListItem in collListItem)
                {

                    DataRow dr = dt.NewRow();
                    foreach (var f in oListItem.FieldValues)
                    {
                        if (j < count)
                        {
                            if (f.Value != null)
                            {
                                value = f.Value.ToString();
                                if (f.Value.ToString() == "Microsoft.SharePoint.Client.FieldUserValue")
                                {
                                    string fname = f.Key.ToString();
                                    if (oListItem[fname] != null)
                                    {
                                        FieldUserValue fuv = (FieldUserValue)oListItem[fname];

                                        value = fuv.Email.ToString();


                                    }
                                }
                                if (f.Value.ToString() == "Microsoft.SharePoint.Client.FieldUserValue[]")
                                {
                                    string fname = f.Key.ToString();
                                    if (oListItem[fname] != null)
                                    {
                                        foreach (FieldUserValue userValue in oListItem[fname] as FieldUserValue[])
                                        {

                                            value = userValue.Email.ToString();

                                        }




                                    }
                                }
                            }

                            key = f.Key.ToString();

                            dr[key] = value;

                            value = "";
                            j = j + 1;

                        }

                    }
                    dt.Rows.Add(dr);
                    j = 0;


                }
                if (selValue.ToUpper() == "EXCEL")
                {
                    DownloadData(listname);
                }
                else
                {
                    SaveToDb(listname);
                }
            }
            catch (Exception e)
            {
                lblRes.Text = e.Message;
                lblRes.ForeColor = Color.Red;
            }


        }

        public void DownloadData(string listname)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Data");
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=" + listname + ".xlsx");
                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }

        }
        public void SaveToDb(string listname)
        {
            try
            {
                DataTable data = dt;

                listname = listname.Replace(' ', '_');
                int n;
                tableStructure = "IF OBJECT_ID('DBO." + listname + "') IS NOT NULL BEGIN" +
                                  " DROP TABLE DBO." + listname + "" +
                  " CREATE TABLE DBO." + listname + "(" +
                  "#coloumns#" +
                  ")" +
                  " END  " +
                  "ELSE " +
                  "BEGIN " +
                  " CREATE TABLE DBO." + listname + "(" +
                  "#coloumns#" +
                  ")" +
                  "END";
                string colStructure = string.Empty;

                for (int k = 0; k < data.Columns.Count; k++)
                {
                    int z;
                    bool isNumeric = int.TryParse(data.Rows[0][k].ToString(), out z);
                    DateTime dtime;
                    bool isDate = DateTime.TryParse(data.Rows[0][k].ToString(), out dtime);
                    Boolean b;
                    bool isbool = Boolean.TryParse(data.Rows[0][k].ToString(), out b);
                    if (isNumeric)
                    {
                        colStructure = colStructure + " [" + data.Columns[k].ColumnName.ToString() + "] int ,";

                    }
                    else if (isDate)
                    {
                        colStructure = colStructure + " [" + data.Columns[k].ColumnName.ToString() + "] datetime ,";

                    }
                    else if (isbool)
                    {
                        colStructure = colStructure + " [" + data.Columns[k].ColumnName.ToString() + "] bit ,";
                    }
                    else
                    {
                        colStructure = colStructure + " [" + data.Columns[k].ColumnName.ToString() + "] nvarchar(500) ,";

                    }
                }
                colStructure = colStructure.TrimEnd(',');
                tableStructure = tableStructure.Replace("#coloumns#", colStructure);
                string connectionString = string.Empty;
                if(ddlSelectAuth.SelectedValue.ToUpper().Contains("WINDOWS"))
                { 
                 connectionString = @"Data Source=#ds#;Initial Catalog=#db#;Integrated Security=SSPI;";
                }
                else
                {
                    connectionString = @"Data Source=#ds#;Database=#db#;User ID=#ID#;Password=#PASS#;Min Pool Size=2;max pool size=20;connection timeout=240";
                }
                //string connectionString = @"Data Source=dehensvbivm055\app_sqlserver_1;Database=CDLdb;User ID=CDLdbnew;Password=cdldb;Min Pool Size=2;max pool size=20;connection timeout=240;Application Name=Henkel CDL";
                connectionString = connectionString.Replace("#ds#", txtDataSource.Text.Trim());
                connectionString = connectionString.Replace("#db#", txtDataBase.Text.Trim());
                connectionString = connectionString.Replace("#ID#", txtUserName.Text.Trim());
                connectionString = connectionString.Replace("#PASS#", txtPassword.Text.Trim());
                bool istablecreated = CreateTableInDB(tableStructure, connectionString);
                if (istablecreated)
                {
                    BulkInsert(connectionString, listname);
                    lblRes.Text = "List item details are successfully save to DBO." + listname + " table";
                    lblRes.ForeColor = Color.Green;
                }
            }
            catch (Exception e)
            {
                if (e.Message.Contains("Could not open a connection to SQL Server") || e.Message.Contains("A network-related or instance-specific error occurred while establishing a connection to SQL Server"))
                {

                    lblRes.Text = "Please enter the correct database details or check your write permissions on the database";
                    lblRes.ForeColor = Color.Red;

                }
                else
                {
                    lblRes.Text = e.Message;
                    lblRes.ForeColor = Color.Red;
                }
            }

        }
        public void BulkInsert(string con, string listname)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(con))
                {
                    cn.Open();
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(cn))
                    {
                        bulkCopy.DestinationTableName = "dbo." + listname + "";
                        bulkCopy.WriteToServer(dt);
                    }
                    cn.Close();
                }
            }
            catch (Exception e)
            {
                lblRes.Text = e.Message;
                lblRes.ForeColor = Color.Red;
            }
        }
        public bool CreateTableInDB(string script, string con1)
        {
            try
            {

                string cmdText = string.Empty;
                cmdText = script.ToString();
                SqlConnection con = new SqlConnection(con1);
                SqlCommand cmd = new SqlCommand();

                cmd.Connection = con;
                cmd.CommandText = cmdText;
                cmd.CommandType = CommandType.Text;



                con.Open();
                return (cmd.ExecuteNonQuery() != 0);
                con.Close();
            }
            catch (SqlException ex)
            {
                //Tracelog.WriteTrace("SOURCE: " + this.GetType().Name + "." + MethodBase.GetCurrentMethod().Name + ". Exception: " + ex.Message, Tracelog.Autherrorseverity.Exception);
                throw;
            }
            catch (Exception ex)
            {
                //Tracelog.WriteTrace("SOURCE: " + this.GetType().Name + "." + MethodBase.GetCurrentMethod().Name + ". Exception: " + ex.Message, Tracelog.Autherrorseverity.Exception);
                throw;
            }
        }


        public void createTable(string[] col, string selValue, string listname)
        {
            try
            {
                dt = new DataTable();

                for (int i = 0; i < col.Length; i++)
                {
                    dt.Columns.Add(new DataColumn(col[i].ToString(), typeof(string)));
                }

            }
            catch (Exception e)
            {
                lblRes.Text = e.Message;
                lblRes.ForeColor = Color.Red;
            }

        }

        protected void brnExport_Click(object sender, EventArgs e)
        {
            //---------------------------------------------------
            if (DropDownList1.SelectedValue.ToUpper() == "EXCEL")
            {

                string k = "";
                for (int i = 0; i < chksharePointGroups.Items.Count; i++)
                {
                    if (chksharePointGroups.Items[i].Selected)
                    {

                        k = k + chksharePointGroups.Items[i].Text + ",";
                    }

                }
                string lsits = k.TrimEnd(',');
                string[] lst = lsits.Split(',');
                string selValue = DropDownList1.SelectedValue;
               
                    for (int i = 0; i < lst.Length; i++)
                    {

                        ReadData(lst[i], selValue);
                    }
                
            }
            //------------------------------------------------------
            else
            {
                bool isValidate = false;
                if (ddlSelectAuth.SelectedValue.ToUpper() == "WINDOWS")
                {
                    if (txtDataBase.Text == "" || txtDataSource.Text == "")
                    {
                        lblRes.Text = "Please enter datasource and database";
                        lblRes.ForeColor = Color.Red;
                    }
                    else
                    {
                        isValidate = true;
                    }
                }
                else if (ddlSelectAuth.SelectedValue.ToUpper().Contains("SQL"))
                {
                    if (txtDataBase.Text == "" || txtDataSource.Text != "" || txtUserName.Text == "" || txtPassword.Text == "")
                    {
                        lblRes.Text = "Please enter dabase details";
                        lblRes.ForeColor = Color.Red;
                    }
                    else
                    {
                        isValidate = true;
                    }
                }
                else
                {
                    isValidate = true;
                }

                if(isValidate)
                {
                    string z = "";
                    for (int i = 0; i < chksharePointGroups.Items.Count; i++)
                    {
                        if (chksharePointGroups.Items[i].Selected)
                        {

                            z = z + chksharePointGroups.Items[i].Text + ",";
                        }

                    }
                    string lsitss = z.TrimEnd(',');
                    string[] lsts = lsitss.Split(',');
                    string selValues = DropDownList1.SelectedValue;
                    
                        for (int i = 0; i < lsts.Length; i++)
                        {

                            ReadData(lsts[i], selValues);
                        }
                    
                   
                }
            }
        }



    }
}