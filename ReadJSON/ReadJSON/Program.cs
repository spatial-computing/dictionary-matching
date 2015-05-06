using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.OleDb;
using Newtonsoft.Json.Linq;
using System.Data;
using Excel=Microsoft.Office.Interop.Excel;
using Npgsql;


namespace ReadJSON
{
    class Program
    {
        static void Main(string[] args)
        {
            //For a file which is given in the format specified in UpdateResults.xls, update the DictionaryResult column with the new values from the JSON files.
            UpdateResults();

            //Reads JSON results into a Excel file
            //readJSOnfile();

            //For a file which is given in the format specified in UpdateResults.xls, find the precision, recall and F score
            //findPerf();

            //Update the Excel column to check if the GroundTruth exists in dictionary
           //CheckDictionary();

            //Postgres SQL update in table (incomplete)
            //updateTable();
           
        }
    
        private static void UpdateResults()
        {
            Excel._Application oApp = new Excel.Application();
            oApp.Visible = true;

            Excel.Workbook oWorkbook = oApp.Workbooks.Open("C:\\Users\\rmenon\\1920-test-maps\\UpdateDictResults.xls");
            Excel.Worksheet oWorksheet = oWorkbook.Worksheets["Sheet1"];

            int rowNo = oWorksheet.UsedRange.Rows.Count-3;
            object[,] array = oWorksheet.UsedRange.Value;
  
            for (int filenum = 1; filenum <= 10; filenum++)
            {
                string path = @"C:\Users\rmenon\1920-test-maps\1920-" + filenum + ".pngByPixels.txt";
                if (File.Exists(path))
                {
                    string jsonstring = File.ReadAllText(path);
                    jsonstring = jsonstring.Replace(",}", "}").Replace("},]", "}]");
                    jsonstring = jsonstring.Substring(0, jsonstring.Length - 2);
                    JObject rss = JObject.Parse(jsonstring);
                    for (int i = 0; i < rss["features"].Count(); i++)
                    {
                        string nameBeforeDictionary = (string)rss["features"][i]["NameBeforeDictionary"];
                        string nameAfterDictionary = (string)rss["features"][i]["NameAfterDictionary"];
                        string sameMatches = (string)rss["features"][i]["SameMatches"];
                        //results.WriteLine(filenum.ToString() + '\t' + nameBeforeDictionary + '\t' + nameAfterDictionary);

                            for(int j=2; j<=rowNo; j++)
                            {
                                if(array[j,1].ToString()==filenum.ToString() && array[j,2].ToString()==nameBeforeDictionary)
                                {
                                    array[j, 3] = nameAfterDictionary;
                                    array[j, 8] = sameMatches;
                                }
                            }
                    }

                }
            }

            oWorksheet.UsedRange.Value = array;

            oWorkbook.Save();
            oWorkbook.Close();
            oApp.Quit();

            oWorksheet = null;
            oWorkbook = null;
            oApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

        }

        private static void updateTable()
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=169.254.95.120;Port=5432;User Id=joe;Password=postgres;Database=postgres;");
            conn.Open();

            NpgsqlCommand command = new NpgsqlCommand("COPY myCopyTestTable FROM STDIN", conn);
            NpgsqlCopyIn cin = new NpgsqlCopyIn(command, conn, Console.OpenStandardInput()); // expecting input in server encoding!
            try
            {
                cin.Start();
            }
            catch (Exception e)
            {
                try
                {
                    cin.Cancel("Undo copy");
                }
                catch (NpgsqlException e2)
                {
                    // we should get an error in response to our cancel request:
                    if (!("" + e2).Contains("Undo copy"))
                    {
                        throw new Exception("Failed to cancel copy: " + e2 + " upon failure: " + e);
                    }
                }
                throw e;
            }

            conn.Close();
        }

        private static void CheckDictionary()
        {
         
            string dictpath = Directory.GetCurrentDirectory() + @"\dict_all.txt";
            string content = File.ReadAllText(dictpath);
            string[] words = content.Split('\n');

            Excel._Application oApp = new Excel.Application();
            oApp.Visible = true;

            Excel.Workbook oWorkbook = oApp.Workbooks.Open("C:\\Users\\rmenon\\1920-test-maps\\Datasets_modified.xlsx");
            Excel.Worksheet oWorksheet = oWorkbook.Worksheets["Sheet1"];

            int rowNo = oWorksheet.UsedRange.Rows.Count;
            object[,] array = oWorksheet.UsedRange.Value;

        
            if (array[1,4].ToString() == "MapName")
            {
                for(int i=2; i<=rowNo; i++)
                {
                    string mapname = (string)array[i, 4];
                    if(Array.FindAll(words, s => s.Equals(mapname.ToLower())).Count()>=1)
                    {
                        array[i, 5] = 1;
                    }
                    else
                    {
                        array[i, 5] = 0;
                    }
                }
            }
            oWorksheet.UsedRange.Value = array;     

            oWorkbook.Save();
            oWorkbook.Close();
            oApp.Quit();

            oWorksheet = null;
            oWorkbook = null;
            oApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }

        private static void readJSOnfile()
        {
            string resultspath = @"C:\Users\rmenon\1920-test-maps\results_new.xls";
            File.WriteAllText(resultspath, string.Empty);
            StreamWriter results = new StreamWriter(resultspath, append: true);
            results.AutoFlush = true;
            results.WriteLine("FileNumber\tNameBeforeDictionary\tNameAfterDictionary");
            for (int filenum = 1; filenum <= 10; filenum++)
            {
                string path = @"C:\Users\rmenon\1920-test-maps\1920-" + filenum + ".pngByPixels.txt";
                if (File.Exists(path))
                {
                    string jsonstring = File.ReadAllText(path);
                    jsonstring = jsonstring.Replace(",}", "}").Replace("},]", "}]");
                    jsonstring = jsonstring.Substring(0, jsonstring.Length - 2);
                    JObject rss = JObject.Parse(jsonstring);
                    for (int i = 0; i < rss["features"].Count(); i++)
                    {
                        string nameBeforeDictionary = (string)rss["features"][i]["NameBeforeDictionary"];
                        string nameAfterDictionary = (string)rss["features"][i]["NameAfterDictionary"];
                        results.WriteLine(filenum.ToString() + '\t' + nameBeforeDictionary + '\t' + nameAfterDictionary);


                    }

                }
            }
        }

        private static void findPerf()
        {
            OleDbConnection MyConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\\Users\\rmenon\\1920-test-maps\\UpdateResults_NW.xls';Extended Properties=Excel 8.0;HDR=Yes;IMEX=1;");
            OleDbDataAdapter MyCommand = new OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
            //MyCommand.TableMappings.Add("Table", "TestTable");
            DataSet dataSet = new DataSet();
            MyCommand.Fill(dataSet);
            DataTable myTable = dataSet.Tables[0];
            myTable.CaseSensitive = false;


            DataRow[] preciserows = myTable.Select("OCRResult IS NOT NULL AND OCRResult = MapName");
            DataRow[] totalrows = myTable.Select("OCRResult IS NOT NULL");
            float precision = ((float)preciserows.Count() / (float)totalrows.Count()) * 100;
            Console.WriteLine("Precision of OCR results = {0} %", precision);

            DataRow[] recallrows = myTable.Select("OCRResult = MapName");
            totalrows = myTable.Select();
            float recall = ((float)recallrows.Count() / (float)totalrows.Count()) * 100;
            Console.WriteLine("Recall of OCR results = {0} %\n", recall);

            float FScore = (2 * precision * recall) / (precision + recall);
            Console.WriteLine("F-Score of OCR results = {0} %\n", FScore);

            preciserows = myTable.Select("DictionaryResult IS NOT NULL AND DictionaryResult = MapName");
            totalrows = myTable.Select("DictionaryResult IS NOT NULL");
             precision = ((float)preciserows.Count() / (float)totalrows.Count()) * 100;
            Console.WriteLine("Precision after dictionary matching = {0} %", precision);

            recallrows = myTable.Select("DictionaryResult = MapName");
            totalrows = myTable.Select();
            recall = ((float)recallrows.Count() / (float)totalrows.Count()) * 100;
            Console.WriteLine("Recall after dictionary matching {0} %\n", recall);

             FScore = 2 * precision * recall / (precision + recall);
             Console.WriteLine("F-Score of dictionary matching = {0} %\n", FScore);

             Console.WriteLine("*** Evaluating the results using OCRResult if DictionaryResult is empty ***\n");

             preciserows = myTable.Select("ModifiedDictionaryResult IS NOT NULL AND ModifiedDictionaryResult = MapName");
             totalrows = myTable.Select("ModifiedDictionaryResult IS NOT NULL");
             precision = ((float)preciserows.Count() / (float)totalrows.Count()) * 100;
             Console.WriteLine("Precision for modified dictionary results = {0} %", precision);

             recallrows = myTable.Select("ModifiedDictionaryResult = MapName");
             totalrows = myTable.Select();
             recall = ((float)recallrows.Count() / (float)totalrows.Count()) * 100;
             Console.WriteLine("Recall for modified dictionary results {0} %\n", recall);

             FScore = 2 * precision * recall / (precision + recall);
             Console.WriteLine("F-Score of modified dictionary results = {0} %\n\n", FScore);

            Console.WriteLine("Press ENTER to return...");
            Console.ReadLine();
        }
    }
}
