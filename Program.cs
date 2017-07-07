using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.IO;


namespace OCR
{
    class Program
    {
        static void testowe()
        {
            //           //testowe
            //           int count = 1;
            //           String[] dictionary = { "Artikelcode", "EAN", "Nr.", "Ihre Art.Nr", "Article", "Nr. IhreArt.Nr" }; //Pobierane powinno byc z bazy danych

            //           Console.WriteLine("Nazwa pliku {0}", Path.GetFileName(@"C:\test\ex1.xls"));

            //           Regex pattern;

            //           int indexOfEnd = 0;
            //           string[] patterns = { "[a-zA-Z0-9]{2}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{2}-[a-zA-Z0-9]{1,4}",  //Product code
            //                                 "[a-zA-Z0-9]",                //Product description
            //                                 "^[0-9]{1,},00$",             //Quantity
            //                                 "^[0-9]{1,},[0-9]{1,2}$",     // Price
            //                                 "^[0-9]{1,},[0-9]{1,2}$",     //Total Value == Price
            //                               };
            //           //Dla kodow
            //           //    string[] patternTEST = { "[a-zA-Z0-9]{2}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{2}-[a-zA-Z0-9]{1,4}$", "^[0-9]{3}-[0-9]{10}$", "^[0-9]-[0-9]{4}-[0-9]{2}-[0-9]{2}$", "^[A-Z][0-9]{2}[A-Z]{3}$", "^[a-zA-Z0-0]{1,}/[a-zA-Z0-0]{1,}$" };
            //           string[] patternTEST = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Code'", "Pattern");

            //           // 0: codes   1: description    2: quantity   3: price  4: Discount 5: TotalValue
            ////Dodac total value, discounts

            //           var rows = new List<Row>();


            //           Regex regex2 = new Regex(@"^[a-zA-Z0-0]{1,}/[a-zA-Z0-0]{1,}$");
            //           Regex regex = new Regex(patterns[0]);



            //           Console.WriteLine(indexOf);




            //foreach (DataRow dr in ex.ds.Tables[0].Rows)
            //{

            //    for (int i = 0; i < ex.ds.Tables[0].Columns.Count; i++)
            //    {
            //        //Looking for index of column with CODES (using column names)
            //        foreach (string term in dictionary)
            //        {
            //            if (dr[i].Equals(term))
            //            {
            //                Console.WriteLine("Found code column! " + i); //Do usuniecia potem
            //                indexOf = ex.ds.Tables[0].Rows.IndexOf(dr);
            //                Console.WriteLine("Column begining index: {0}\n", indexOf); //Do usuniecia potem

            //                //columns[0] = codes
            //                //for (int k = 0; k < columns.Length; k++) { columns[k] = i + k; } //Tutaj powinno przypissc kolumne z kodami + kolukmne z opisem 
            //                columns[0] = i;
            //                break;
            //            }
            //        }

            //        //Jezeli znajdzie po nazwie to nie szuka juz po patternie i konczy foreach dla wierszy
            //        if (indexOf != 0) { break; }



            //        //w petli przypisuje wartosci regex kolejnym wartosciom z tablicy i sprawdza wiersze
            //        //jezeli znajdzie to przerwya petle i konczy


            //        foreach (string pat in patternTEST)
            //        {
            //            regex = new Regex(pat);

            //            if (regex.IsMatch(dr[i].ToString()))
            //            {
            //                count++;
            //                Console.WriteLine("Teraz zaczynamy part2 dla kodow");

            //                // for (int k = 0; k < columns.Length; k++) { columns[k] = i + k; }
            //                columns[0] = i;
            //                indexOf = ex.ds.Tables[0].Rows.IndexOf(dr) - count;
            //                //    break; //To jeszcze przemyslec. Ale skoro znalazlo pasujacy wiersz to jest to ta kolumna. W dalszej czesci kodu bedzie sprawdzac wszystkie wiersze
            //            }
            //        }

            //    }
            //}








            ///* ************************************** DESCRIPTION *********************************************** */

            //string[] dictionaryDescription = { "Description", "Descripción artículo", "Beschreibung", "Artikelbeschrijving", "D e s i g n a t i o n", "Designation" };


            ////Jezeli mamy juz kody to sprawdza kolumne po prawej
            //if (columns[0] != -1)
            //{
            //    int i = columns[0] + 1;
            //    regex = new Regex(patterns[1]);
            //    foreach (DataRow dr in ex.ds.Tables[0].Rows)
            //    {
            //        foreach (string term in dictionaryDescription)
            //        {
            //            if (dr[i].Equals(term))
            //            {
            //                Console.WriteLine("Found description! " + i); //Do usuniecia potem
            //                // indexOf = ex.ds.Tables[0].Rows.IndexOf(dr);
            //                Console.WriteLine("Column begining index: {0}\n", indexOf); //Do usuniecia potem

            //                //columns[0] = codes
            //                //for (int k = 0; k < columns.Length; k++) { columns[k] = i + k; } //Tutaj powinno przypissc kolumne z kodami + kolukmne z opisem 
            //                columns[1] = i;
            //                break;
            //            }
            //        }
            //    }
            //}



            ///* ************************************** QUANTITY *********************************************** */
            ////getting patterns from sql table
            //patternTEST = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Quantity'", "Pattern");

            ////If description column index is known then checking column to the right TO FIND QUANTITY 
            //if (columns[1] == -1)
            //{
            //    int i = columns[1] + 1;

            //    foreach (string pat in patternTEST)
            //    {
            //        regex = new Regex(pat);
            //        foreach (DataRow dr in ex.ds.Tables[0].Rows)
            //        {
            //            if (regex.IsMatch(dr[i].ToString()) && !(string.IsNullOrEmpty(dr[i].ToString())))
            //            {
            //                count++;
            //                Console.WriteLine("Teraz zaczynamy part2 dla ilosci {0}", dr[i].ToString());

            //                // for (int k = 0; k < columns.Length; k++) { columns[k] = i + k; }
            //                columns[2] = i;
            //            }
            //        }
            //    }
            //}
            //// Looking for Q-ty column in all document
            //else
            //{
            //    bool found = false;

            //    foreach (string pat in patternTEST)
            //    {
            //        regex = new Regex(pat);

            //        foreach (DataRow dr in ex.ds.Tables[0].Rows)
            //        {
            //            for (int i = 0; i < ex.ds.Tables[0].Columns.Count; i++)
            //            {
            //                // Line is not empty && not found yet && pattern matches
            //                if (!found && !(string.IsNullOrEmpty(dr[i].ToString())) && regex.IsMatch(dr[i].ToString()))
            //                {
            //                    columns[2] = i;
            //                    Console.WriteLine("test {0}", columns[2]);
            //                    found = true;
            //                }
            //                else { continue; }
            //            }
            //        }
            //    }
            //}


            ////Tutaj w peti dla qty i price - ALE TO TYLKO BEDZIE TUTAJ DZIALAC



            // Console.WriteLine("Columns[1] = {0} PO SZUKANIU. Index {1}", columns[1], indexOf);












            //********************************************************************DISPLAYING COLUMNS******************************************************************

            //////Wyswietlanie wierszy i pobieranie
            //foreach (DataRow dr in ex.ds.Tables[0].Rows)
            //{
            //    //Printing rows > index of Header
            //    if (ex.ds.Tables[0].Rows.IndexOf(dr) > indexOf)
            //    {
            //        //Sprawdzenie czy linia nie jest pusta
            //        if (!(string.IsNullOrEmpty(dr[0].ToString())) || !(string.IsNullOrEmpty(dr[1].ToString())))
            //        {

            //            //Sprawdzenie czy linia pasuje do wzorca
            //            for (int i = 0; i < columns.Length; i++)
            //            {
            //                pattern = new Regex(patterns[i]);

            //                if (pattern.IsMatch(dr[i].ToString()) || regex2.IsMatch(dr[i].ToString()))
            //                {

            //                    //Jesli tak to dodaje do listy pasujacych elementow
            //                    // kody.Add(dr[i].ToString());
            //                    //Console.Write(dr[i].ToString() + "\t\t" + ex.ds.Tables[0].Rows.IndexOf(dr));


            //                    if (i == 0)
            //                    {
            //                        //Powinno stworzyc tyle obiektow ile wierszy,
            //                        //Numery kolumn maja byc pobierane z tablicy columns[]
            //                        r = new Row();
            //                        r.Code = dr[columns[0]].ToString();
            //                        // r.Description = dr[i + 1].ToString();
            //                        r.Description = dr[columns[1]].ToString();
            //                        r.Quantity = dr[2].ToString();
            //                        r.Price = dr[3].ToString();
            //                        rows.Add(r);

            //                    }
            //                }
            //            }
            //        }
            //        else
            //        {
            //            Console.WriteLine("\nEnd of column index: {0}", indexOfEnd = ex.ds.Tables[0].Rows.IndexOf(dr));
            //            break;
            //        }
            //    }
            //}




        }

        public static string getItemCode(string row)
        {
            string match = "S/Ref.: ";
            int indexOfString = Regex.Match(row.ToString(), match).Index + match.Length;
            row = row.ToString().Substring(indexOfString, (row.ToString().Length - indexOfString));
            bool hasSpace = row.Contains(" ");

            if (row.Length > 15 && !hasSpace)
            {
                return row.Substring(0, 15);
            }
            else if (row.Length > 15 && hasSpace)
            {
                var spaceIndex = row.IndexOf(" ");
                return row.Substring(0, spaceIndex);
            }
            else
            {
                return row.Substring(0, row.Length);
            }
        }

        public static double toDouble(string num)
        {
            try
            {
                if (string.IsNullOrEmpty(num))
                    return 0;

                int posComa = num.LastIndexOf(",");
                int posPunto = num.LastIndexOf(".");

                if (posComa != -1) num = (num.Replace(".", "").Replace(",", "."));//.Replace(",", ".")
                else if (posComa < posPunto) num = (num.Replace(",", ""));//.Replace(",", ".")

                //   else num = num.Replace(",", "").Replace(".", ",");


                return double.Parse(num);
            }
            catch (Exception e)
            {
                return 0;
                throw new Exception("toDouble > " + e.Message);
               
            }
        }

        static string[] getDataFromSQL(string query, string sqlColumnName)
        {
            string temp;
            List<string> listPaths = new List<string>();

            SqlConnection connection = new SqlConnection("Data Source=10.0.0.8;Initial Catalog=dev001;User ID=sebastian;Password=proc2017");
            connection.Open();

            SqlCommand command = new SqlCommand(query, connection);

            SqlDataReader reader = command.ExecuteReader();


            while (reader.Read())
            {
                temp = reader.GetString(reader.GetOrdinal(sqlColumnName));
                listPaths.Add(temp);
            }

            reader.Close();
            connection.Close();

            return listPaths.ToArray();
        }

        static void insertDataToSQL(List<Row> rows, string docName)
        {

            SqlConnection connection = new SqlConnection("Data Source=10.0.0.8;Initial Catalog=dev001;User ID=sebastian;Password=proc2017");
            string insertCommand = "INSERT INTO [dbo].[Comandes_lines] ([item] ,[item_codigo_prediccion],[cantidad], [cantidad_prev],[precio_bruto] ,[precioBruto_prev],[descuento],[importe], [docName]) " + " VALUES (@code, @code, @quantity, @quantity,  @price, @price, @discount, @totalValue, @docName)";



            connection.Open();
            SqlTransaction trans = connection.BeginTransaction();

            SqlCommand cmd = new SqlCommand(insertCommand, connection, trans);

            cmd.Parameters.Add("@code", DbType.String);
            cmd.Parameters.Add("@quantity", DbType.Double);
            cmd.Parameters.Add("@price", DbType.Double);
            cmd.Parameters.Add("@discount", DbType.String);
            cmd.Parameters.Add("@totalValue", DbType.Double);
            cmd.Parameters.Add("@docName", DbType.String);


            foreach (var item in rows)
            {
                try
                {
                    cmd.Parameters[0].Value = item.Code;
                    cmd.Parameters[1].Value = item.Quantity;
                    cmd.Parameters[2].Value = item.Price;
                    cmd.Parameters[3].Value = item.Discount;
                    cmd.Parameters[4].Value = item.TotalValue;
                    cmd.Parameters[5].Value = docName;

                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    continue;
                }

            }

            trans.Commit();
            connection.Close();

        }

        static void findColumnsByName(MS_excel excel, ref int[] columns, ref int indexOf)
        {
            string[] dictionary = { "" };
            bool found = false;
            bool indexOfBegining = false;
            string[] columnNames = { "Code", "Description", "Quantity", "Price", "Discount", "TotalValue" };


            for (int i = columns[0]; i < columns.Length; i++)
            {
                found = false;

                switch (i)
                {
                    case 0:
                        dictionary = getDataFromSQL("SELECT * FROM Dictionary WHERE ColumnName = 'Code'", "Term");
                        break;
                    case 1:
                        dictionary = getDataFromSQL("SELECT * FROM Dictionary WHERE ColumnName = 'Description'", "Term");
                        break;
                    case 2:
                        dictionary = getDataFromSQL("SELECT * FROM Dictionary WHERE ColumnName = 'Quantity'", "Term");
                        break;
                    case 3:
                        dictionary = getDataFromSQL("SELECT * FROM Dictionary WHERE ColumnName = 'Price'", "Term");
                        break;
                    case 4:
                        dictionary = getDataFromSQL("SELECT * FROM Dictionary WHERE ColumnName = 'Discount'", "Term");
                        break;
                    case 5:
                        dictionary = getDataFromSQL("SELECT * FROM Dictionary WHERE ColumnName = 'TotalValue'", "Term");
                        break;
                    default:
                        Array.Clear(dictionary, 0, dictionary.Length);
                        break;
                }

                foreach (DataRow dr in excel.ds.Tables[0].Rows)
                {

                    for (int k = 0; k < excel.ds.Tables[0].Columns.Count; k++)
                    {
                        foreach (string term in dictionary)
                        {
                            if (!found && dr[k].Equals(term))
                            {
                                if (!indexOfBegining)
                                {
                                    indexOf = indexOf = excel.ds.Tables[0].Rows.IndexOf(dr);
                                    indexOfBegining = true;
                                }

                                Console.WriteLine("Found column " + columnNames[i] + " on position {0}", k); //Do usuniecia potem
                                columns[i] = k;
                                found = true;
                            }
                            else { continue; }
                        }
                    }

                }
            }
        }

        static bool checkPattern(MS_excel excel, DataRow dr, string[] patterns, int indexOf, int[] columns, int i)
        {
            Regex regex;
            bool match = false;

            if (excel.ds.Tables[0].Rows.IndexOf(dr) > indexOf && (!(string.IsNullOrEmpty(dr[0].ToString())) || !(string.IsNullOrEmpty(dr[1].ToString()))))
            {
                foreach (string pat in patterns)
                {
                    if (columns[i] == -1) { break; }
                    regex = new Regex(pat);
                    if (regex.IsMatch(dr[columns[i]].ToString()))
                    {
                        match = true;
                    }
                }
            }


            if (match) { return true; }
            else { return false; }

        }

        //Special value = true should be used for SONEPAR's invoices
        static void displayRows(MS_excel excel, int[] columns, int indexOf, string file, bool sonepar = false)
        {

            Row r;
            var rows = new List<Row>();
            string docName = file;
            bool matchCode, matchDescription, matchQuantity, matchPrice, matchDiscount, matchTotalValue;
            matchCode = false;
            matchDescription = false;
            matchQuantity = false;
            matchPrice = false;
            matchDiscount = false;
            matchTotalValue = false;


            string[] pattCode = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Code'", "Pattern");
            string[] pattDescription = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Description'", "Pattern");
            string[] pattQuantity = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Quantity'", "Pattern");
            string[] pattPrice = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Price'", "Pattern");
            string[] pattDiscount = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Discount'", "Pattern");
            string[] pattTotalValue = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'TotalValue'", "Pattern");


            foreach (DataRow dr in excel.ds.Tables[0].Rows)
            {
                for (int i = 0; i < columns.Length; i++)
                {
                    switch (i)
                    {
                        case 0:
                            matchCode = checkPattern(excel, dr, pattCode, indexOf, columns, i);
                            break;
                        case 1:
                            matchDescription = checkPattern(excel, dr, pattDescription, indexOf, columns, i);
                            break;
                        case 2:
                            matchQuantity = checkPattern(excel, dr, pattQuantity, indexOf, columns, i);
                            break;
                        case 3:
                            matchPrice = checkPattern(excel, dr, pattPrice, indexOf, columns, i);
                            break;
                        case 4:
                            matchDiscount = checkPattern(excel, dr, pattDiscount, indexOf, columns, i);
                            break;
                        case 5:
                            matchTotalValue = checkPattern(excel, dr, pattTotalValue, indexOf, columns, i);
                            break;
                        default:
                            break;
                    }
                }


                if (sonepar && !String.IsNullOrEmpty(dr[columns[2]].ToString()) && !String.IsNullOrEmpty(dr[columns[3]].ToString()) && excel.ds.Tables[0].Rows.IndexOf(dr) > indexOf)
                {
                    string match = "S/Ref.: ";
                    bool found = false;
                    int counter = 1;
                    string row;
                    r = new Row();

                    if (Regex.IsMatch(dr[columns[1]].ToString(), match))
                    {
                        row = dr[columns[1]].ToString();
                        r.Code = getItemCode(row);

                    }
                    else
                    {
                        string rowQuantity = excel.ds.Tables[0].Rows[excel.ds.Tables[0].Rows.IndexOf(dr) + counter][columns[2]].ToString();
                        string rowPrice = excel.ds.Tables[0].Rows[excel.ds.Tables[0].Rows.IndexOf(dr) + counter][columns[3]].ToString();

                        while (!found && String.IsNullOrEmpty(rowQuantity.ToString()) && String.IsNullOrEmpty(rowPrice.ToString()))
                        {
                            row = excel.ds.Tables[0].Rows[excel.ds.Tables[0].Rows.IndexOf(dr) + counter][columns[1]].ToString();

                            if (Regex.IsMatch(row.ToString(), match))
                            {
                                r.Code = getItemCode(row);
                                found = true;
                            }
                            else { counter++; }
                        }
                    }

                    try
                    {
                        if (columns[1] != -1) r.Description = dr[columns[1]].ToString(); else { r.Description = ""; }
                        if (columns[2] != -1) r.Quantity = Convert.ToDouble(dr[columns[2]].ToString().Replace(",", "."));
                        if (columns[3] != -1 && matchPrice) r.Price = toDouble(dr[columns[3]].ToString());
                        if (columns[4] != -1) r.Discount = dr[columns[4]].ToString(); else { r.Discount = ""; }
                        if (columns[5] != -1 && matchTotalValue) r.TotalValue = toDouble(dr[columns[5]].ToString());
                        rows.Add(r);
                    }
                    catch { continue; }





                }
                else if ((matchPrice && matchQuantity) || (matchQuantity && matchTotalValue))
                {

                    r = new Row();
                    try
                    {
                        if (columns[0] != -1 && !String.IsNullOrEmpty(dr[columns[0]].ToString())) { r.Code = dr[columns[0]].ToString(); } else { r.Code = ""; }

                        if (columns[1] != -1) r.Description = dr[columns[1]].ToString();
                        if (columns[2] != -1 && matchQuantity) r.Quantity = Convert.ToDouble(dr[columns[2]].ToString().Replace(",", "."));
                        if (columns[3] != -1 && matchPrice) r.Price = toDouble(dr[columns[3]].ToString());
                        if (columns[4] != -1) r.Discount = dr[columns[4]].ToString(); else { r.Discount = ""; }
                        if (columns[5] != -1 && matchTotalValue) { r.TotalValue = toDouble(dr[columns[5]].ToString()); }
                        Console.Write("Q {0} {2}  P {1} {3} \n", matchQuantity, matchPrice, dr[columns[2]], dr[columns[3]]);
                        rows.Add(r);

                    }

                    catch
                    {
                    //    rows.Add(r);
                        continue;
                    }
                }

                matchCode = false;
                matchDescription = false;
                matchQuantity = false;
                matchPrice = false;
                matchDiscount = false;
                matchTotalValue = false;

            }

            insertDataToSQL(rows, docName);
        }



        static void Main(string[] args)
        {
            /*
             * COLUMNS contains ID of columns
             * columns[0] = Codes
             * columns[1] = Description
             * columns[2] = Qty
             * columns[3] = Price
             * columns[4] = Discount
             * columns[5] = TotalValue of row
             */
            ////directory to the folder with files
            string directory = @"C:\test\Sonepar2\";


            ////Creating list of files to be read
            var dir = new DirectoryInfo(directory);
            FileInfo[] fi = dir.GetFiles("*.xls");

            foreach (FileInfo file in fi)
            {
                string pathToFile = String.Concat(directory, file);
                bool sonepar = false;

                if (file.Name.ToString().Contains("sonepar")) { sonepar = true; }

                try
                {

                    MS_excel ex = new MS_excel(@pathToFile, "Sheet1", false);

                    //Adding values -1 to the columns array
                    int indexOf = 0;
                    int[] columns = new int[ex.ds.Tables[0].Columns.Count];
                    for (int i = 0; i < ex.ds.Tables[0].Columns.Count - 1; i++) { columns[i] = -1; }

                    Console.WriteLine(pathToFile);

                    findColumnsByName(ex, ref columns, ref indexOf);


                    displayRows(ex, columns, indexOf, file.ToString(), sonepar);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    continue;
                }
            }


            //string[] pattCode = getDataFromSQL("SELECT * FROM Patterns WHERE ColumnName = 'Code'", "Pattern");
            //MS_excel excel = new MS_excel(@"C:\[OCR] hotFolders\sebastian2\done\TEst\0000242004_kod w opisie.xls", "Sheet1", false);
            //foreach (DataRow dr in  excel.ds.Tables[0].Rows) 
            //{
            //    string pattern = "(" + "[a-zA-Z0-9]{2}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{3}" + ")";
            //    string test = "PX-0024-OXI SIROS E27 IP23 100W";
            //    Match mtch = Regex.Match(test, pattern);
            //    //Console.WriteLine(pattern);
            //    if (mtch.Success) Console.WriteLine(mtch.Groups[1].Value);


            //}




        }
    }
}
