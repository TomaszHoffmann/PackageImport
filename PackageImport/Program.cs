using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using System;
using System.Text;
using System.Data.SqlClient;

namespace PackageImport
{
    class Program
    {
                static void LoadExcelMain(string excelPath)
        {

            string conString = string.Empty;
            string extension = Path.GetExtension(excelPath);
       
            switch (extension)
            {
                case ".xls": //Excel 97-03
                    conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx": //Excel 07 or higher
                    conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                    break;

            }
    
            conString = string.Format(conString, excelPath);
            using (OleDbConnection excel_con = new OleDbConnection(conString))
            {
                excel_con.Open();
                //read sheet named TABLE_NAME from excell , with columns : Name , Salary
                string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                DataTable dtExcelData = new DataTable();

                //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                dtExcelData.Columns.AddRange(new DataColumn[3] { new DataColumn("Id", typeof(int)),
            new DataColumn("Name", typeof(string)),
            new DataColumn("Salary",typeof(decimal)) });

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                {
                    oda.Fill(dtExcelData);
                }
                excel_con.Close();
             
                string consString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                using (SqlConnection con = new SqlConnection(consString))
                using (SqlCommand command = con.CreateCommand())
                {
                    con.Open();
                    command.Parameters.Add("@NR_ART", SqlDbType.NVarChar);
                    command.Parameters.Add("@JM1", SqlDbType.Int);
                    foreach (DataRow row in dtExcelData.Rows)
                    {
                        // string test = Convert.ToString(row["Materiał"]);
                        int jm1;
                        try { jm1=Convert.ToInt32(row["Ilo#opak#1"]);
                        }
                        catch (Exception) { jm1 = 0; }

                        command.CommandText = "UPDATE DBO.NHANDLOTABLE SET opak1=@JM1 WHERE NR_ART=@NR_ART";

                        // command.Parameters.AddWithValue("@NR_ART", Convert.ToString(row["Materiał"]));
                        
                        command.Parameters["@NR_ART"].Value = Convert.ToString(row["Materiał"]);


                        command.Parameters["@JM1"].Value = jm1;

                        command.ExecuteNonQuery();
                        Console.WriteLine(Convert.ToString(row["Materiał"]) + " complete");
                    }
                    con.Close();
                }


            }
        }
   
        
        static void Main(string[] args)
        {
            LoadExcelMain(@"C:\Users\tomasz.hoffmann\Desktop\pt.xlsx");
        }
    }
}
