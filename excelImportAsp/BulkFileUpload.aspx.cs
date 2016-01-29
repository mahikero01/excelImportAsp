using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Configuration;

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

        private void ExcelBulkUpload() 
        {
            String strFileNmae = "StudentRecords.xlsx";
            String strSheetName = "StudentDetails";
            String path = Server.MapPath("Import") + "\\" + strFileNmae;
            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;HDR=Yes;IMEX=2";
        
            OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
            OleDbCommand cmd = new OleDbCommand("Select * from [" + strSheetName + "$]", excelConnection);
            excelConnection.Open();

            OleDbDataReader dReader;
            dReader = cmd.ExecuteReader();

            string strServer = ConfigurationManager.ConnectionStrings["StudentConnectionString"].ToString();

            SqlBulkCopy sqlBulk = new SqlBulkCopy(strServer);
            


        }
    }
}