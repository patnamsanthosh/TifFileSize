using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TifFileSize
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            string con = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\santh\Downloads\TestFileSize.xlsx; Extended Properties = 'Excel 12.0 Xml;HDR=YES;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {

                OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "select * from [Sheet1$]";
                    comm.Connection = connection;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);

                    }

                }
            }

            if (dt.Rows.Count > 0)
            {
                DataTable workTable = new DataTable("TIFfiles");

                workTable.Columns.Add("RequestID", typeof(String));
                workTable.Columns.Add("FileName", typeof(String));
                workTable.Columns.Add("FilePath", typeof(String));
                workTable.Columns.Add("FileSize", typeof(String));
                workTable.Columns.Add("ChartNavId", typeof(int));
                // List<int> vs = new List<int>();
                foreach (DataRow item in dt.AsEnumerable())
                {
                    string[] files = Directory.GetFiles(item["FilePath"].ToString());
                    foreach (var file in files)
                    {
                        if (Path.GetFileName(file) == (string)item["FileName"])
                        {
                           
                            DataRow newCustomersRow = workTable.NewRow();

                            newCustomersRow["RequestID"] = item["RequestID"];
                            newCustomersRow["FileName"] = item["FileName"];
                            newCustomersRow["FilePath"] = item["FilePath"];
                            newCustomersRow["FileSize"] = BytesToString(new FileInfo(file).Length);
                            newCustomersRow["ChartNavId"] = item["ChartNavId"];
                            workTable.Rows.Add(newCustomersRow);
                        }
                    }

                }

                StringBuilder sb = new StringBuilder();
                //adding header
                sb.Append("RequestID,FileName,FilePath,FileSize,ChartNavId");
                sb.AppendLine();
                foreach (DataRow dr in workTable.Rows)
                {
                    foreach (DataColumn dc in workTable.Columns)
                        sb.Append(FormatCSV(dr[dc.ColumnName].ToString()) + ",");
                    sb.Remove(sb.Length - 1, 1);
                    sb.AppendLine();
                }
                File.WriteAllText("C:\\Santhosh\\tfiresult.csv", sb.ToString());
            }
            else
            {
                Console.WriteLine("No records found");
            }
            //DataSet dsResult = new DataSet();
            //dsResult.Tables.Add(workTable);
            // ExportDataSet(dsResult, "C:\\Santhosh\\tfiresult.xlsx");

        }

        static String BytesToString(long byteCount)
        {
            string[] suf = { "B", "KB", "MB", "GB", "TB", "PB", "EB" }; //Longs run out around EB
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString() + suf[place];
        }

        public static string FormatCSV(string input)
        {
            try
            {
                if (input == null)
                    return string.Empty;

                bool containsQuote = false;
                bool containsComma = false;
                int len = input.Length;
                for (int i = 0; i < len && (containsComma == false || containsQuote == false); i++)
                {
                    char ch = input[i];
                    if (ch == '"')
                        containsQuote = true;
                    else if (ch == ',')
                        containsComma = true;
                }

                if (containsQuote && containsComma)
                    input = input.Replace("\"", "\"\"");

                if (containsComma)
                    return "\"" + input + "\"";
                else
                    return input;
            }
            catch
            {
                throw;
            }
        }

        
    }
}
