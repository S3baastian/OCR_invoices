using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace OCR
{
    public class MS_excel
    {
        public DataSet ds = new DataSet();
        public bool valid = false, read = false;
        public string path, clientId, tabMark;
        //public Order order;
        private bool exactTab = false;
        private int row_lines_start = 22, tabPos = 0;

        //Konstruktory sciezka, sheetmark=? exactMark=?
        public MS_excel(string filepath, string sheetMark, bool exactMark)
        {
            path = filepath;
            tabMark = sheetMark;
            exactTab = exactMark;

            read_excel_oledb();

           // validacion();
        }
        //Konstruktor: sciezka i pozycja arkusza = numeracja zaczyna sie od 1 w MS Excel
        public MS_excel(string filepath, int sheetPos)
        {
            this.path = filepath;
            tabPos = sheetPos;
            if (tabPos == 0) tabPos = 1;
            read_excel_oledb();
        }


        //Sprawdza czy plik to excel
        public static bool isExcel(string file)
        {
            try
            {
                if (file.ToLower().EndsWith(".xls") || file.ToLower().EndsWith(".xlsx"))
                    return true;
            }
            catch (Exception e)
            {
                //func_aux.reportError("", 0, "UUE", "MS_excel.isExcel", e.Message); 
            }
            return false;
        }

        private void read_excel_oledb()
        {
            try
            {
                //Polaczenie sie z plikiem excela
                var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" +
                    "Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";

                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = tabQuery_set(sheets);
                        var adapter = new OleDbDataAdapter(cmd);
                        adapter.Fill(ds);
                        try { 
                            if (ds.Tables[0].Rows.Count > 0) read = true; 
                        }
                        catch { }
                    }
                }
                
            }
            catch (Exception e) { 
              //  func_aux.reportError("", 0, "UUE", "read_excel_oledb(path=" + path + ")", e.Message); 
                Console.WriteLine(e);
            }
           
        }

        //Tutaj pobiera dane z arkusza, wiersz po wierszu
        private string tabQuery_set(DataTable sheets)
        {
            try
            {
                if (sheets.Rows.Count > 1)
                {
                    if (!string.IsNullOrEmpty(tabMark))
                    {////Busqueda por nombre del tab
                        foreach (DataRow dr in sheets.Rows)
                        {
                            string tabName = dr["TABLE_NAME"].ToString().ToLower();

                            if (exactTab)
                            {
                                tabName = tabName.Replace("$", "");
                                tabName = tabName.Replace("'", "");
                                tabName = tabName.Trim();
                            }

                            if (!exactTab && tabName.Contains(tabMark.ToLower()))
                            {
                                return "SELECT * FROM [" + dr["TABLE_NAME"].ToString() + "] ";
                            }
                            else if (exactTab && tabName.Equals(tabMark.ToLower()))
                            {
                                return "SELECT * FROM [" + dr["TABLE_NAME"].ToString() + "] ";
                            }
                        }
                    }
                    else if (tabPos >= 1 && sheets.Rows.Count >= tabPos)
                    {
                        return "SELECT * FROM [" + sheets.Rows[tabPos - 1]["TABLE_NAME"].ToString() + "] ";
                    }
                }
                else
                    return "SELECT * FROM [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] ";

                throw new Exception("tab search failed");
            }
            catch (Exception e)
            {
                //func_aux.reportError("", 0, "UUE", "tabQuery_set(path=" + path + ")", e.Message);
                throw new Exception("tabQuery_set > " + e.Message);
            }
        }
    }
}
