using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics ;
using System.Globalization;
using System.Drawing;

namespace ExcelXML
{
    public static class XML
    {
        /// <summary>
        /// Excel table settings
        /// </summary>
        public static class Table
        {
            /// <summary>
            /// Set rows height (default = 0) (as System.Integer)
            /// </summary>
            private static int rowsHeight = 0;

            /// <summary>
            /// Set table text color (as System.Drawing.Color)
            /// </summary>
            private static System.Drawing.Color textColor = System.Drawing.Color.Black;

            /// <summary>
            /// Set table text to upper (as System.Boolean)
            /// </summary>
            private static bool toUpper = false;

            /// <summary>
            /// Defines font of table
            /// </summary>
            private static System.Drawing.Font tableFont = new System.Drawing.Font("Calibri", 12);

            /// <summary>
            /// Set rows auto size (as System.Bool)
            /// </summary>
            private static bool rowAutoHeight = true;

            /// <summary>
            /// Reduces font size if cell text lenght is more than this number (default = 0 (do not reduce)) (as System.Integer)
            /// </summary>
            private static int reductFontSize = 0;

            /// <summary>
            /// Set if table has a filter
            /// </summary>
            private static bool hasFilter = false;

            /// <summary>
            /// Set number format to text for each table
            /// </summary>
            private static bool setTextFormat = false;

            public static int RowsHeight { get => rowsHeight; set => rowsHeight = value; }
            public static Color TextColor { get => textColor; set => textColor = value; }
            public static bool ToUpper { get => toUpper; set => toUpper = value; }
            public static Font TableFont { get => tableFont; set => tableFont = value; }
            public static bool RowAutoHeight { get => rowAutoHeight; set => rowAutoHeight = value; }
            public static int ReductFontSize { get => reductFontSize; set => reductFontSize = value; }
            public static bool HasFilter { get => hasFilter; set => hasFilter = value; }
            public static bool SetTextFormat { get => setTextFormat; set => setTextFormat = value; }
        }

        /// <summary>
        /// Excel table header settings
        /// </summary>
        public static class Header
        {
            /// <summary>
            /// Set columns width (default = 0) (as System.Integer)
            /// </summary>
            private static int columnsWidth = 0;

            /// <summary>
            /// Set header row height (default = 25) (as System.Integer)
            /// </summary>
            private static int rowsHeight = 0;

            /// <summary>
            /// Set header text color (as System.Drawing.Color)
            /// </summary>
            private static System.Drawing.Color textColor = System.Drawing.Color.Black;

            /// <summary>
            /// Defines font of header
            /// </summary>
            private static System.Drawing.Font headerFont = new System.Drawing.Font("Calibri", 12);

            /// <summary>
            /// Set header text to upper (as System.Boolean)
            /// </summary>
            private static bool toUpper = false;

            /// <summary>
            /// Set columns autosize (as System.Bool)
            /// </summary>
            private static bool columnsAutoWidth = true;

            /// <summary>
            /// Set rows auto size (as System.Bool)
            /// </summary>
            private static bool rowAutoHeight = true;

            /// <summary>
            /// Set rows auto size (as System.Bool)
            /// </summary>
            private static bool showHeader = true;

            public static int ColumnsWidth { get => columnsWidth; set => columnsWidth = value; }
            public static int RowsHeight { get => rowsHeight; set => rowsHeight = value; }
            public static Color TextColor { get => textColor; set => textColor = value; }
            public static Font HeaderFont { get => headerFont; set => headerFont = value; }
            public static bool ToUpper { get => toUpper; set => toUpper = value; }
            public static bool ColumnsAutoWidth { get => columnsAutoWidth; set => columnsAutoWidth = value; }
            public static bool RowAutoHeight { get => rowAutoHeight; set => rowAutoHeight = value; }
            public static bool ShowHeader { get => showHeader; set => showHeader = value; }
                        
            ///// <summary>
            ///// Reduces font size if cell text lenght is more than this number (default = 0 (do not reduce)) (as System.Integer)
            ///// </summary>
            //public static int ReductFontSize = 0;
        }

        /// <summary>
        /// Printing settings
        /// </summary>
        public static class Printing
        {
            /// <summary>
            /// Set print orientation (as System.Boolean)
            /// </summary>
            private static bool landscape = false;

            /// <summary>
            /// Set paper type false = A4; true = A3 (default = false)  (as System.Boolean)
            /// </summary>
            private static bool a3 = false;

            public static bool Landscape { get => landscape; set => landscape = value; }
            public static bool A3 { get => a3; set => a3 = value; }
        }

        /// <summary>
        /// Output file properties settings
        /// </summary>
        public static class FileProperties
        {
            /// <summary>
            /// Set author name (as System.String)
            /// </summary>
            private static string authorName = "System";

            /// <summary>
            /// Set company name (as System.String)
            /// </summary>
            private static string companyName = "";

            public static string AuthorName { get => authorName; set => authorName = value; }
            public static string CompanyName { get => companyName; set => companyName = value; }
        }

        /// <summary>
        /// Worksheet settings
        /// </summary>
        public static class Worksheet
        {
            /// <summary>
            /// Set sheet name (as System.String)
            /// </summary>
            private static string name = "List";

            public static string Name { get => name; set => name = value; }
        }

        private static string flName = "";

        /// <summary>
        /// Remove diacritics = true; reserve diacritics = false (default = false) (as System.Boolean)
        /// </summary>
        private static bool withoutDiacritics = false;

        public static bool WithoutDiacritics { get => withoutDiacritics; set => withoutDiacritics = value; }
        public static string FlName { get => flName; set => flName = value; }

        /// <summary>
        /// Save DataTable as XML (Excel table), More sheets
        /// </summary>
        /// <param name="dtData">DataTable[] for export to Excel file (as System.Data.DataTable)</param>
        /// <param name="tableName">Name of Excel table (as System.String)</param>
        /// <param name="fileName">Full path and file name (as System.String)</param>
        public static void Create(DataTable[] dtDataArray, string fileName, string tableName = "")
        {
            if (fileName == "")
            {
                System.Windows.Forms.SaveFileDialog sfDialog = new SaveFileDialog();
                sfDialog.ShowDialog();
                fileName = sfDialog.FileName;
            }

            FlName = fileName;

            try
            {
                if (File.Exists(fileName) == true) File.Delete(fileName);
            }

            catch { Debug.WriteLine("Failed to delete existing file."); }//System.Windows.Forms.MessageBox.Show("File named " + fileName + " already exists and is in use by another process.\nPlease, kill that process first.", "Caution", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation); return; }

            //Hlavička XML
            xmlWrite("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            xmlWrite("<?mso-application progid=\"Excel.Sheet\"?>");
            xmlWrite("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            xmlWrite("xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            xmlWrite("xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
            xmlWrite("xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            xmlWrite("xmlns:html=\"http://www.w3.org/TR/REC-html40\">");

            xmlWrite("<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">");
            xmlWrite("<Author>" + FileProperties.AuthorName + "</Author>");
            xmlWrite("<LastAuthor>" + FileProperties.AuthorName + "</LastAuthor>");


            xmlWrite("<Created>" + DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ") + "</Created>");
            xmlWrite("<LastSaved>" + DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ") + "</LastSaved>");
            xmlWrite("<Company>" + FileProperties.CompanyName + "</Company>");
            xmlWrite("<Version>12.00</Version>");
            xmlWrite("</DocumentProperties>");

            xmlWrite("<ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">");
            xmlWrite("<WindowHeight>5190</WindowHeight>");
            xmlWrite("<WindowWidth>18195</WindowWidth>");
            xmlWrite("<WindowTopX>120</WindowTopX>");
            xmlWrite("<WindowTopY>135</WindowTopY>");
            xmlWrite("<ProtectStructure>False</ProtectStructure>");
            xmlWrite("<ProtectWindows>False</ProtectWindows>");
            xmlWrite("</ExcelWorkbook>");

            xmlWrite("<Styles>");
            xmlWrite("<Style ss:ID=\"Default\" ss:Name=\"Normal\">");
            xmlWrite("<Alignment ss:Vertical=\"Bottom\"/>");
            xmlWrite("<Borders/>");
            xmlWrite("<Font ss:FontName=\"Calibri\" x:CharSet=\"238\" x:Family=\"Swiss\" ss:Size=\"11\"");
            xmlWrite("ss:Color=\"#000000\"/>");
            xmlWrite("<Interior/>");
            xmlWrite("<NumberFormat/>");
            xmlWrite("<Protection/>");
            xmlWrite("</Style>");

            xmlWrite("<Style ss:ID=\"s15\">");
            xmlWrite("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>");
            xmlWrite("<Borders>");
            xmlWrite("<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("</Borders>");
            xmlWrite("<Font ss:FontName=\"" + Table.TableFont.FontFamily.Name + "\" x:CharSet=\"238\" x:Family=\"Swiss\" ss:Size=\"" + Math.Round(Table.TableFont.Size, 0).ToString() + "\"");
            xmlWrite("ss:Color=\"" + System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(Table.TextColor.ToArgb())) + "\"/>");
            if (Table.SetTextFormat == true)
                xmlWrite("<NumberFormat ss:Format=\"@\"/>");
            else
                xmlWrite("<NumberFormat ss:Format=\"General\"/>");
            xmlWrite("</Style>");

            xmlWrite("<Style ss:ID=\"s16\">");
            xmlWrite("<NumberFormat ss:Format=\"@\"/>");
            xmlWrite("</Style>");

            xmlWrite("<Style ss:ID=\"s17\">");
            xmlWrite("<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>");
            xmlWrite("<Borders>");
            xmlWrite("<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("</Borders>");
            xmlWrite("<Font ss:FontName=\"Calibri\" x:CharSet=\"238\" x:Family=\"Swiss\" ss:Size=\"7.5\"");
            xmlWrite("ss:Color=\"#000000\"/>");
            xmlWrite("<NumberFormat ss:Format=\"@\"/>");
            xmlWrite("</Style>");

            xmlWrite("<Style ss:ID=\"s18\">");
            xmlWrite("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>");
            xmlWrite("<Borders>");
            xmlWrite("<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            xmlWrite("</Borders>");
            xmlWrite("<Font ss:FontName=\"" + Table.TableFont.FontFamily.Name + "\" x:CharSet=\"238\" x:Family=\"Swiss\" ss:Size=\"" + Math.Round(Table.TableFont.Size / 2, 0).ToString() + "\"");
            xmlWrite("ss:Color=\"" + System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(Table.TextColor.ToArgb())) + "\"/>");
            if (Table.SetTextFormat == true)
                xmlWrite("<NumberFormat ss:Format=\"@\"/>");
            else
                xmlWrite("<NumberFormat ss:Format=\"General\"/>");
            xmlWrite("</Style>");

            xmlWrite("<Style ss:ID=\"s19\">");
            xmlWrite("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>");
            xmlWrite("<Borders>");
            xmlWrite("<Border ss:Position=\"Bottom\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("<Border ss:Position=\"Left\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("<Border ss:Position=\"Right\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("<Border ss:Position=\"Top\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("</Borders>");
            xmlWrite("<Font ss:FontName=\"" + Header.HeaderFont.FontFamily.Name + "\" x:CharSet=\"238\" x:Family=\"Swiss\" ss:Size=\"" + Header.HeaderFont.Size + "\"");
            xmlWrite("ss:Color=\"" + System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(Header.TextColor.ToArgb())) + "\" ss:Bold=\"1\"/>");
            xmlWrite("<NumberFormat ss:Format=\"@\"/>");
            xmlWrite("</Style>");

            xmlWrite("<Style ss:ID=\"s20\">");
            xmlWrite("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>");
            xmlWrite("<Borders>");
            xmlWrite("<Border ss:Position=\"Bottom\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("<Border ss:Position=\"Left\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("<Border ss:Position=\"Right\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("<Border ss:Position=\"Top\" ss:LineStyle=\"Double\" ss:Weight=\"3\"/>");
            xmlWrite("</Borders>");
            xmlWrite("<Font ss:FontName=\"Calibri\" x:CharSet=\"238\" x:Family=\"Swiss\" ss:Size=\"7.5\"");
            xmlWrite("ss:Color=\"#000000\" ss:Bold=\"1\"/>");
            xmlWrite("<NumberFormat ss:Format=\"@\"/>");
            xmlWrite("</Style>");

            xmlWrite("</Styles>");

            int sheetNum = 1;
            foreach (DataTable dtData in dtDataArray)
            {
                if (dtData.TableName.Length > 0)
                    if (Worksheet.Name == dtData.TableName)
                        Worksheet.Name = Worksheet.Name + sheetNum.ToString();
                    else
                        Worksheet.Name = dtData.TableName;
                else Worksheet.Name = Worksheet.Name + sheetNum.ToString();

                xmlWrite("<Worksheet ss:Name=\"" + Worksheet.Name + "\">");
                xmlWrite("<Names>");
                xmlWrite("<NamedRange ss:Name=\"Print_Titles\" ss:RefersTo=\"=" + Worksheet.Name + "!R1\"/>");
                xmlWrite("</Names>");

                xmlWrite("<Table ss:ExpandedColumnCount=\"" + (dtData.Columns.Count + 1).ToString() + "\" ss:ExpandedRowCount=\"" + (dtData.Rows.Count + 1).ToString() + "\" x:FullColumns=\"1\"");

                xmlWrite("x:FullRows=\"0\" ss:DefaultRowHeight=\"15\">");

                //set column width
                foreach (DataColumn dcData in dtData.Columns)
                {

                    if (Header.ColumnsWidth == 0)
                    {
                        double charCount = 0;
                        foreach (DataRow drData in dtData.Rows)
                        {

                            int filterSize = 0;
                            if (Table.HasFilter == true) filterSize = 20;

                            if (Math.Round(dcData.ColumnName.Length * (Table.TableFont.Size * 0.58) + Header.HeaderFont.Size + filterSize, 0) > charCount)
                            {
                                charCount = Math.Round(dcData.ColumnName.Length * (Table.TableFont.Size * 0.58) + Header.HeaderFont.Size + filterSize, 0);
                            }
                            else
                            {
                                if (Table.ReductFontSize == 0)
                                {
                                    if (Math.Round(drData[dcData.ColumnName].ToString().Length * (Table.TableFont.Size * 0.58) + Table.TableFont.Size, 0) > charCount)
                                        charCount = Math.Round(drData[dcData.ColumnName].ToString().Length * (Table.TableFont.Size * 0.58) + Table.TableFont.Size, 0);
                                }
                                else
                                {
                                    if (Math.Round(drData[dcData.ColumnName].ToString().Length * (Table.ReductFontSize * 0.58) + Table.ReductFontSize, 0) > charCount)
                                        charCount = Math.Round(drData[dcData.ColumnName].ToString().Length * (Table.ReductFontSize * 0.58) + Table.ReductFontSize, 0);
                                }

                            }

                        }
                        xmlWrite("<Column ss:AutoFitWidth=\"1\" ss:Width=\"" + charCount.ToString() + "\"/>");
                    }
                    else xmlWrite("<Column ss:AutoFitWidth=\"1\" ss:Width=\"" + Header.ColumnsWidth.ToString() + "\"/>");

                }

                //Table header
                if (Header.ShowHeader == true)
                {
                    if (Header.RowsHeight == 0)
                    {
                        double rowHeight = Header.HeaderFont.Size * (1.3);
                        xmlWrite("<Row ss:Height=\"" + Math.Round(rowHeight, 0).ToString() + "\">");
                    }
                    else
                        xmlWrite("<Row ss:Height=\"" + Header.RowsHeight.ToString() + "\">");

                    foreach (DataColumn dcData in dtData.Columns)
                    {
                        if (Header.ToUpper == true) dcData.ColumnName = dcData.ColumnName.ToUpper();
                        xmlWrite("<Cell ss:StyleID=\"s19\"><Data ss:Type=\"String\">" + dcData.ColumnName + "</Data><NamedCell");
                        if (Table.HasFilter == true) xmlWrite("ss:Name=\"_FilterDatabase\"/><NamedCell ss:Name=\"Print_Titles\"/></Cell>");
                        else xmlWrite("ss:Name=\"Print_Titles\"/></Cell>");
                    }
                    xmlWrite("</Row>");
                }

                //Table body
                foreach (System.Data.DataRow drData in dtData.Rows)
                {
                    //Set rows height
                    if (Table.RowsHeight == 0)
                    {
                        double rowHeight = Table.TableFont.Size * (1.3);
                        xmlWrite("<Row ss:Height=\"" + Math.Round(rowHeight, 0).ToString() + "\">");
                    }
                    else
                        xmlWrite("<Row ss:Height=\"" + Table.RowsHeight.ToString() + "\">");

                    //Insert rows
                    foreach (DataColumn dcData in dtData.Columns)
                    {
                        string cellText = removeDia(drData[dcData.ColumnName].ToString());

                        string numberFormat = "String";
                        if (Table.SetTextFormat == false)
                        {
                            numberFormat = "Number";
                            try { double.Parse(cellText); cellText = cellText.Replace(",", "."); }
                            catch { numberFormat = "String"; }
                        }

                        if (removeDia(drData[dcData.ColumnName].ToString()).Length > Table.ReductFontSize)
                            if (Table.ReductFontSize > 0)
                            {
                                if (Table.ToUpper == true) cellText = cellText.ToUpper();
                                xmlWrite("<Cell ss:StyleID=\"s18\"><Data ss:Type=\"" + numberFormat + "\">" + cellText + "</Data></Cell>");
                            }
                            else
                                xmlWrite("<Cell ss:StyleID=\"s15\"><Data ss:Type=\"" + numberFormat + "\">" + cellText + "</Data></Cell>");
                        else
                            xmlWrite("<Cell ss:StyleID=\"s15\"><Data ss:Type=\"" + numberFormat + "\">" + cellText + "</Data></Cell>");
                    }
                    xmlWrite("</Row>");
                }

                xmlWrite("</Table>");

                xmlWrite("<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">");
                xmlWrite("<PageSetup>");
                if (Printing.Landscape == true) xmlWrite("<Layout x:Orientation=\"Landscape\"/>");
                xmlWrite("<Header x:Margin=\"0.3\" x:Data=\"&amp;L" + tableName + "&amp;R" + DateTime.Now.Date.ToString().Replace(" 0:00:00", "") + "\"/>");
                xmlWrite("<Footer x:Margin=\"0.3\" x:Data=\"&amp;RStrana &amp;P / &amp;N\"/>");
                xmlWrite("<PageMargins x:Bottom=\"0.78740157499999996\" x:Left=\"0.7\" x:Right=\"0.7\"");
                xmlWrite("x:Top=\"0.78740157499999996\"/>");
                xmlWrite("</PageSetup>");
                xmlWrite("<Print>");
                xmlWrite("<ValidPrinterInfo/>");
                if (Printing.A3 == true) xmlWrite("<PaperSizeIndex>8</PaperSizeIndex>");
                xmlWrite("<HorizontalResolution>600</HorizontalResolution>");
                xmlWrite("<VerticalResolution>600</VerticalResolution>");
                xmlWrite("</Print>");
                xmlWrite("<Selected/>");
                xmlWrite("<ProtectObjects>False</ProtectObjects>");
                xmlWrite("<ProtectScenarios>False</ProtectScenarios>");
                xmlWrite("</WorksheetOptions>");
                if (Table.HasFilter == true)
                {
                    xmlWrite("<AutoFilter x:Range=\"R1C1:R1C" + dtData.Columns.Count.ToString() + "\"");
                    xmlWrite("xmlns=\"urn:schemas-microsoft-com:office:excel\">");
                    xmlWrite("</AutoFilter>");
                }
                xmlWrite("</Worksheet>");
                sheetNum++;
            }

            xmlWrite("</Workbook>");

        }

        private static void xmlWrite(string element)
        {
            Encoding enc = Encoding.GetEncoding("Windows-1250");
            try
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(FlName, true, enc);
                file.WriteLine(element);
                file.Close();
            }
            catch { throw; }//System.Windows.Forms.MessageBox.Show("Soubor s názvem " + "PC_" + flName + " již existuje a je používán jiným procesem.\n\nUkončete, prosím, tento proces.", "Pozor", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation); return; }
        }

        private static string removeDia(string text)
        {
            byte[] tempBytes;
            tempBytes = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(text);
            return System.Text.Encoding.UTF8.GetString(tempBytes);
        }
    }

    public static class CSV
    {
        private static string flName = "";

        /// <summary>
        /// Select separator char. (Default = ;) (as System.String)
        /// </summary>
        private static string separator = ";";

        /// <summary>
        /// Remove diacritics = true; reserve diacritics = false (default = false) (as System.Boolean)
        /// </summary>
        private static bool withoutDiacritics = false;

        /// <summary>
        /// Append csv to existing file (default = false) (as System.Boolean)
        /// </summary>
        private static bool append = false;

        /// <summary>
        /// Write header to csv file (default = true) (as System.Boolean)
        /// </summary>
        private static bool writeHeader = true;

        public static string FlName { get => flName; set => flName = value; }
        public static string Separator { get => separator; set => separator = value; }
        public static bool WithoutDiacritics { get => withoutDiacritics; set => withoutDiacritics = value; }
        public static bool Append { get => append; set => append = value; }
        public static bool WriteHeader { get => writeHeader; set => writeHeader = value; }

        public static void Create(DataTable dtData, string fileName)
        {
            if (fileName.Length == 0) { throw new Exception("Filename can't be null."); }
            FlName = fileName;

            if (Append == false)
                try { File.Delete(FlName); }
                catch (Exception ex) { }

            //CSV header
            int colNum = 0;
            string rowToWrite = "";
            foreach (DataColumn dc in dtData.Columns)
            {
                colNum++;
                string strToWrite = dc.ColumnName;
                if (colNum != dtData.Columns.Count) strToWrite += ";";
                rowToWrite += strToWrite;
            }

            if (WriteHeader == true) csvWrite(removeDia(rowToWrite));

            //CSV body
            foreach (DataRow dr in dtData.Rows)
            {
                rowToWrite = "";
                colNum = 0;
                foreach (DataColumn dc in dtData.Columns)
                {
                    colNum++;
                    string strToWrite = dr[dc.ColumnName].ToString();
                    if (colNum != dtData.Columns.Count) strToWrite += ";";
                    rowToWrite += strToWrite;
                }
                csvWrite(removeDia(rowToWrite));
            }
        }

        public static string CreateString(DataTable dtData)
        {
            string strFile = "";

            //CSV header
            int colNum = 0;
            string rowToWrite = "";
            foreach (DataColumn dc in dtData.Columns)
            {
                colNum++;
                string strToWrite = dc.ColumnName;
                if (colNum != dtData.Columns.Count) strToWrite += ";";
                rowToWrite += strToWrite;
            }

            if (WriteHeader == true) strFile = removeDia(rowToWrite);

            //CSV body
            foreach (DataRow dr in dtData.Rows)
            {
                strFile += "\r\n";

                rowToWrite = "";
                colNum = 0;
                foreach (DataColumn dc in dtData.Columns)
                {
                    colNum++;
                    string strToWrite = dr[dc.ColumnName].ToString();
                    if (colNum != dtData.Columns.Count) strToWrite += ";";
                    rowToWrite += strToWrite;
                }
                strFile += removeDia(rowToWrite);
            }

            return strFile;
        }

        private static void csvWrite(string element)
        {
            //Encoding enc = Encoding.GetEncoding("Windows-1250");
            try
            {
                //System.IO.StreamWriter file = new System.IO.StreamWriter(flName, true, enc);
                File.AppendAllText(FlName,element + Environment.NewLine);
                //file.WriteLine(element);
                //file.Close();
            }
            catch (Exception ex){ }
        }

        private static string removeDia(string text)
        {
            if (WithoutDiacritics == true)
            {
                var normalizedString = text.Normalize(NormalizationForm.FormD);
                var stringBuilder = new StringBuilder();

                foreach (var c in normalizedString)
                {
                    var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                    if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                        stringBuilder.Append(c);
                }
                return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
            }
            return text;
        }

        /// <summary>
        /// Create DataTable from CSV file
        /// </summary>
        /// <param name="fileName">Path to CSV file</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable(string fileName)
        {
            DataTable dtReturn = new DataTable("csv");

            string strFile = File.ReadAllText(fileName);
            string[] row = strFile.Split('\n');

            //csv header
            string header = row[0];
            foreach (string columnName in header.Split(Convert.ToChar(Separator)))
            {
                dtReturn.Columns.Add(columnName.Replace("\r", "").Replace("\n", ""));
            }

            //csv data
            int rowCount = row.Length;
            for (int rowNum = 1; rowNum < rowCount; rowNum++)
            {
                DataRow drReturn = dtReturn.NewRow();
                string[] rowData = row[rowNum].Split(Convert.ToChar(Separator));

                int colNum = 0;
                foreach (string data in rowData)
                {
                    if (colNum > dtReturn.Columns.Count - 1)
                    {
                        break;
                        //dtReturn.Columns.Add("Column" + (colNum + 1).ToString());
                    }
                    else
                    {
                        drReturn[colNum] = data.Replace("\r", "").Replace("\n", "");
                        colNum++;
                    }
                }

                if (drReturn[0].ToString().Length > 0)
                    dtReturn.Rows.Add(drReturn);
            }

            return dtReturn;
        }

        /// <summary>
        /// Create DataTable from CSV formatted string
        /// </summary>
        /// <param name="strFile">CSV formatted string</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTableFromString(string strFile)
        {
            DataTable dtReturn = new DataTable("csv");
            
            string[] row = strFile.Split('\n');

            //csv header
            string header = row[0];
            foreach (string columnName in header.Split(Convert.ToChar(Separator)))
            {
                dtReturn.Columns.Add(columnName.Replace("\r", "").Replace("\n", ""));
            }

            //csv data
            int rowCount = row.Length;
            for (int rowNum = 1; rowNum < rowCount; rowNum++)
            {
                DataRow drReturn = dtReturn.NewRow();
                string[] rowData = row[rowNum].Split(Convert.ToChar(Separator));

                int colNum = 0;
                foreach (string data in rowData)
                {
                    if (data.Length > 0 && colNum < dtReturn.Columns.Count)
                        drReturn[colNum] = data.Replace("\r", "").Replace("\n", "");
                    colNum++;
                }

                if (drReturn[0].ToString().Length > 0)
                    dtReturn.Rows.Add(drReturn);
            }

            return dtReturn;
        }
    
    }

    public static class Common
    {
        /// <summary>
        /// Join tables on column value; other columns should have different names.
        /// </summary>
        /// <param name="dtLeft">Left table (master)</param>
        /// <param name="dtRight">Right table (slave)</param>
        /// <param name="joinField">Column name (must be the same in both tables)</param>
        /// <returns>DataTable</returns>
        public static DataTable LeftOuterJoin(DataTable dtLeft, DataTable dtRight, string joinField)
        {
            DataTable dtRes = new DataTable();

            foreach (DataColumn dcLeft in dtLeft.Columns)
                dtRes.Columns.Add(dcLeft.ColumnName);

            foreach (DataColumn dcRight in dtRight.Columns)
                try
                {
                    dtRes.Columns.Add(dcRight.ColumnName);
                }
                catch { }

            foreach (DataRow drLeft in dtLeft.Rows)
            {
                DataRow drRes = dtRes.NewRow();

                foreach (DataRow drRight in dtRight.Select(joinField + "='" + drLeft[joinField].ToString() + "'"))
                {
                    foreach (DataColumn dcLeft in dtLeft.Columns)
                        drRes[dcLeft.ColumnName] = drLeft[dcLeft.ColumnName].ToString();
                    foreach (DataColumn dcRight in dtRight.Columns)
                        drRes[dcRight.ColumnName] = drRight[dcRight.ColumnName].ToString();

                    dtRes.Rows.Add(drRes);
                }
            }

            return dtRes;

        }

        /// <summary>
        /// TODO!!
        /// </summary>
        /// <param name="dtLeft"></param>
        /// <param name="dtRight"></param>
        /// <param name="joinField"></param>
        /// <returns>DataTable</returns>
        public static DataTable Join(DataTable dtLeft, DataTable dtRight, string joinField)
        {
            DataTable dataTable = new DataTable();
            foreach (DataColumn column in dtLeft.Columns)
                dataTable.Columns.Add(column.ColumnName);
            foreach (DataColumn column in dtRight.Columns)
            {
                try
                {
                    dataTable.Columns.Add(column.ColumnName);
                }
                catch
                {
                }
            }
            foreach (DataRow row1 in dtLeft.Rows)
            {
                DataRow row2 = dataTable.NewRow();
                DataRow[] dataRowArray = dtRight.Select(joinField + "='" + row1[joinField].ToString() + "'");
                if (dataRowArray.Length > 0)
                {
                    foreach (DataRow dataRow in dataRowArray)
                    {
                        foreach (DataColumn column in dtLeft.Columns)
                            row2[column.ColumnName] = row1[column.ColumnName].ToString();
                        foreach (DataColumn column in dtRight.Columns)
                            row2[column.ColumnName] = dataRow[column.ColumnName].ToString();
                        dataTable.Rows.Add(row2);
                    }
                }
                else
                {
                    foreach (DataColumn column in dtRight.Columns)
                        row2[column.ColumnName] = "0";
                    foreach (DataColumn column in dtLeft.Columns)
                        row2[column.ColumnName] = row1[column.ColumnName].ToString();
                    dataTable.Rows.Add(row2);
                }
            }
            return dataTable;
        }
        
        /// <summary>
        /// Order columns in refered DataTable with the same order as columnNames
        /// </summary>
        /// <param name="dtData">Refered DataTable</param>
        /// <param name="columnNames">Names of columns in required order</param>
        /// <param name="removeOthers">FALSE = leave other columns on the end of table, TRUE = remove other columns</param>
        public static void OrderColumns(ref DataTable dtData, string[] columnNames, bool removeOthers = false)
        {
            List<string> stringList = new List<string>();
            foreach (DataColumn column in dtData.Columns)
            {
                bool flag = false;
                foreach (string columnName in columnNames)
                {
                    if (columnName.ToLower() == column.ColumnName.ToLower())
                        flag = true;
                }
                if (!flag)
                    stringList.Add(column.ColumnName);
            }
            if (removeOthers)
            {
                foreach (string name in stringList)
                    dtData.Columns.Remove(name);
            }
            int ordinal = 0;
            foreach (string columnName in columnNames)
            {
                dtData.Columns[columnName].SetOrdinal(ordinal);
                ++ordinal;
            }
        }

        /// <summary>
        /// Returns DataView with ordered columns; similar to OrderColumns function, but do not affects original DataTable
        /// </summary>
        /// <param name="dtData">Original DataTable</param>
        /// <param name="columnNames">Names of columns in required order</param>
        /// <returns>DataView</returns>
        public static DataView ColumnFilter(DataTable dtData, string[] columnNames)
        {
            List<string> stringList = new List<string>();
            foreach (DataColumn column in dtData.Columns)
            {
                bool flag = false;
                foreach (string columnName in columnNames)
                {
                    if (columnName == column.ColumnName)
                        flag = true;
                }
                if (!flag)
                    stringList.Add(column.ColumnName);
            }
            foreach (string name in stringList)
                dtData.Columns.Remove(name);
            int ordinal = 0;
            foreach (string columnName in columnNames)
            {
                dtData.Columns[columnName].SetOrdinal(ordinal);
                ++ordinal;
            }
            return new DataView(dtData);
        }

        /// <summary>
        /// Add or update rows in DataTable
        /// </summary>
        /// <param name="dtData">Original DataTable (update required)</param>
        /// <param name="dtDataNew">DataTable with update data</param>
        /// <param name="primaryKey">Column name to identify the record</param>
        /// <returns>DataTable</returns>
        public static DataTable AddOrUpdate(DataTable dtData, DataTable dtDataNew, string primaryKey)
        {
            DataTable dataTable = dtData.Copy();
            foreach (DataRow row1 in dtDataNew.Rows)
            {
                bool flag = false;
                foreach (DataRow row2 in dataTable.Rows)
                {
                    if (row1[primaryKey].ToString() == row2[primaryKey].ToString())
                    {
                        flag = true;
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            if (column.ColumnName != primaryKey)
                                row2[column.ColumnName] = row1[column.ColumnName];
                        }
                    }
                }
                if (!flag)
                {
                    DataRow row2 = dataTable.NewRow();
                    foreach (DataColumn column in dataTable.Columns)
                        row2[column.ColumnName] = row1[column.ColumnName];
                    dataTable.Rows.Add(row2);
                }
            }
            return dataTable;
        }


    }
}
