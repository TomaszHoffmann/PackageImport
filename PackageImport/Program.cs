using Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace PackageImport
{
    public class Program
    {
        const int NR_ART_LEN = 25;



        DataTable GetTable(DataSet dataset, string name)
        {
            Console.WriteLine("Czytanie arkusza " + name);

            return dataset.Tables[name];

        }

        public class Articles
        {
            public string nr_art;

            public int przelicz, przelskrz, przelpalet;
        }

        public static List<Articles> articles = new List<Articles>();

        bool ParseExcelFile(Stream stream)
        {

            bool failed = false;

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            excelReader.Close();

            DataTable table = GetTable(result, "Arkusz1");

            int count = 0;

            foreach (DataRow row in table.Rows)
            {
                count += 1;
                if (count > 1)
                {
                    string nr_art = Convert.ToString(row[0]);

                    int opak1 = Convert.ToInt32(row[13]);
                    int opak2 = Convert.ToInt32(row[17]);
                    int opak3 = Convert.ToInt32(row[21]);

                   // Console.WriteLine(nr_art + opak1 + opak2 + opak3);

                    Articles art = new Articles() { nr_art = nr_art, przelicz = opak1, przelskrz = opak2, przelpalet = opak3 };

                    articles.Add(art);
                }

            }


            return !failed;
        }

        static void LoadExcelMain(string excelPath, string databseConnection)
        {

            FileStream stream;
            try
            {
                stream = File.Open(excelPath, FileMode.Open, FileAccess.Read);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Import przerwany.");
                Console.ReadKey();
                return;
            }


            bool ok = false;
            Program prog = new Program();

            try
            {
                ok = prog.ParseExcelFile(stream);
                Console.WriteLine("ok");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            string consString = databseConnection;
            using (SqlConnection con = new SqlConnection(consString))
            using (SqlCommand command = con.CreateCommand())
            {
                con.Open();
                command.Parameters.Add("@NR_ART", SqlDbType.NVarChar);
                command.Parameters.Add("@opak1", SqlDbType.Int);
                command.Parameters.Add("@opak2", SqlDbType.Int);
                command.Parameters.Add("@opak3", SqlDbType.Int);

               
           
                foreach (var item in articles)
                {





                    command.CommandText = "UPDATE DBO.NHANDLOTABLE SET przelicz=@opak1, PRZELSKRZ=@opak2, PRZELPALET=@opak3 WHERE NR_ART=@NR_ART";

                    command.Parameters["@NR_ART"].Value = item.nr_art;

                    command.Parameters["@opak1"].Value = item.przelicz; //wiązka
                    command.Parameters["@opak2"].Value = item.przelskrz; //skrzynka
                    command.Parameters["@opak3"].Value = item.przelpalet; //paleta

                    try
                    {
                        command.ExecuteNonQuery();
                        Console.WriteLine(item.nr_art + " complete");
                        
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Console.ReadKey();
                        
                    }

                   


                }
                con.Close();
                Console.WriteLine("Press any key to end");
                Console.ReadKey();

            }





        }

        /*   static void Main(string[] args)
           {
               LoadExcelMain(@"C:\Users\tomasz.hoffmann\Desktop\pt.xlsx", @"Data Source=.\sqlexpress;Initial Catalog=BMW5.6.2.1; Integrated Security=False;User ID=sa;Password=Whokna123@");
           }
        */

        static void Main(string[] args)
        {
            LoadExcelMain(args[0], args[1]);
        }
    }

}
