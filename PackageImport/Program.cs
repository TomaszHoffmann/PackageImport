using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using System;


namespace PackageImport
{
    class Program
    {
                static void LoadExcelMain(string excelPath, string databseConnection)
        {

            string conString = string.Empty;
            string extension = Path.GetExtension(excelPath);
            conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
            conString = string.Format(conString, excelPath);
            using (OleDbConnection excel_con = new OleDbConnection(conString))
            {
                int lp = 0;
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
             
                string consString = databseConnection;
                using (SqlConnection con = new SqlConnection(consString))
                using (SqlCommand command = con.CreateCommand())
                {
                    
                    con.Open();
                    command.Parameters.Add("@NR_ART", SqlDbType.NVarChar);
                    command.Parameters.Add("@opak1", SqlDbType.Int);
                    command.Parameters.Add("@opak2", SqlDbType.Int);
                    command.Parameters.Add("@opak3", SqlDbType.Int);

                   /* command.Parameters.Add("@opisOpak1", SqlDbType.NVarChar);
                    command.Parameters.Add("@opisOpak2", SqlDbType.NVarChar);
                    command.Parameters.Add("@opisOpak3", SqlDbType.NVarChar);*/

                    foreach (DataRow row in dtExcelData.Rows)
                    {
                        lp += 1;
                        string nr_art = Convert.ToString(row["Materiał"]); //nr_art
                       /* string opisOpak1 = Convert.ToString(row["Opis opak1"]);
                        string opisOpak2 = Convert.ToString(row["Opis opak2"]);
                        string opisOpak3 = Convert.ToString(row["Opis opakowanie 3"]);*/

                        int opak1 = Convert.ToInt32(row["Ilo#opak#1"]);
                        int opak2 = Convert.ToInt32(row["Ilo#opak#2"]);
                        int opak3 = Convert.ToInt32(row["Ilo#opak#3"]);

                      

                        command.CommandText = "UPDATE DBO.NHANDLOTABLE SET przelicz=@opak1, PRZELSKRZ=@opak2, PRZELPALET=@opak3 WHERE NR_ART=@NR_ART";

                        command.Parameters["@NR_ART"].Value = nr_art;

                        command.Parameters["@opak1"].Value = opak1; //wiązka
                        command.Parameters["@opak2"].Value = opak2; //skrzynka
                        command.Parameters["@opak3"].Value = opak3; //paleta

                        /*command.Parameters["@opisOpak1"].Value = opisOpak1; 
                        command.Parameters["@opisOpak2"].Value = opisOpak2; 
                        command.Parameters["@opisOpak3"].Value = opisOpak3; */

                        try
                        {
                            command.ExecuteNonQuery();
                            Console.WriteLine(lp+" - "+ nr_art + " complete");
                        }
                        catch(Exception)
                        {
                            Console.WriteLine(lp + " - " + nr_art + " failed");
                        }
                    }
                    con.Close();
                    Console.WriteLine("Press any key to end");
                    Console.ReadKey();
                }


            }
        }
   
        
        static void Main(string[] args)
        {
            LoadExcelMain(@"C:\Users\tomasz.hoffmann\Desktop\pt.xlsx", @"Data Source=.\sqlexpress;Initial Catalog=Test2; Integrated Security=False;User ID=sa;Password=Whokna123@");
        }
    }
}
