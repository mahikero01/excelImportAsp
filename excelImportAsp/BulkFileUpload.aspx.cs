using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Configuration;
using System.Data;
using ClosedXML.Excel;
using System.IO;

namespace excelImportAsp
{
    public partial class BulkFileUpload : System.Web.UI.Page
    {
        String excelConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            ExcelBulkUpload();
        }

        protected void exportExcel(object sender, EventArgs e) {
            
            string strServer = ConfigurationManager.ConnectionStrings["StudentConnectionString"].ToString();

            using (SqlConnection con = new SqlConnection(strServer)) 
            {

                using (SqlCommand cmd = new SqlCommand("SELECT * FROM StudentDetails")) 
                {
                
                    using (SqlDataAdapter sda = new SqlDataAdapter()) 
                    {

                        cmd.Connection = con;
                        sda.SelectCommand = cmd;
                        using (DataTable dt = new DataTable())
                        {

                            sda.Fill(dt);
                            using (XLWorkbook wb = new XLWorkbook()) 
                            {

                                wb.Worksheets.Add(dt, "CustomersRico");

                                Response.Clear();
                                Response.Buffer = true;
                                Response.Charset = "";
                                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                                Response.AddHeader("content-disposition", "attachment; filename=SqlExport.xlsx");

                                using (MemoryStream MyMemoryStream = new MemoryStream())
                                {
                                
                                    wb.SaveAs(MyMemoryStream);
                                    MyMemoryStream.WriteTo(Response.OutputStream);
                                    Response.Flush();
                                    Response.End();


                                }

                            }
                        
                        }

                    }

                }
            
            }

        }

        private void ExcelBulkUpload() 
        {
            if (FileUpload1.HasFile)
            {
                try
                {
                    string path = string.Concat(Server.MapPath("~/Import/" + FileUpload1.FileName));

                    FileUpload1.SaveAs(path);

                    //string strFileNmae = "StudentRecords.xlsx";
                    string strSheetName = "StudentDetails";
                    //String path = Server.MapPath("Import") + "\\" + strFileNmae;
                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + 
                            path + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=2'";

                    OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                    OleDbCommand cmd = new OleDbCommand("Select * from [" + strSheetName + "$]", excelConnection);
                    excelConnection.Open();

                    OleDbDataReader dReader;
                    dReader = cmd.ExecuteReader();

                    string strServer = ConfigurationManager.ConnectionStrings["StudentConnectionString"].ToString();

                    SqlBulkCopy sqlBulk = new SqlBulkCopy(strServer);
                    sqlBulk.DestinationTableName = "StudentDetails";
                    sqlBulk.ColumnMappings.Add("Roll No", "RollNo");
                    sqlBulk.ColumnMappings.Add("Student Name", "Name");
                    sqlBulk.ColumnMappings.Add("Department", "Department");
                    sqlBulk.ColumnMappings.Add("Section", "Section");
                    sqlBulk.WriteToServer(dReader);

                    excelConnection.Close();
                    lbl_Message.Text = "Your file data has been successfully uploaded";
                    lbl_Message.Visible = true;
                }
                catch (Exception ex)
                {
                    lbl_Message.Text = ex.Message;
                }
            }
        }

    }
}