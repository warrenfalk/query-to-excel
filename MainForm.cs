using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace QueryToExcel
{
    public partial class MainForm : Form
    {
        readonly string TemplateFolder = @"C:\Users\wfalk\source\QueryToExcel\Template";
        readonly ConnectionInfo[] Connections = new ConnectionInfo[] {
            new ConnectionInfo {
                Name = "DIV",
                Caption = "GP (DIV)",
                Database = "DIV",
                Server = @"DIVSQL3\GP",
            },
            new ConnectionInfo {
                Name = "Divisions_Inc_MSCRM",
                Caption = "CRM (Divisions_Inc_MSCRM)",
                Database = "Divisions_Inc_MSCRM",
                Server = @"DIVSQL3\CRM",
            },
            new ConnectionInfo {
                Name = "ProviderPortal",
                Caption = "ProviderPortal",
                Database = "ProviderPortal",
                Server = @"DIVSQL5\DATA",
            },
        };
        readonly Dictionary<string, int?> TypeStyleMap = CreateTypeStyleMap();

        private static Dictionary<string, int?> CreateTypeStyleMap()
        {
            Dictionary<string, int?> map = new Dictionary<string, int?>();
            map.Add("int", 0);
            map.Add("smallint", 0);
            map.Add("tinyint", 0);
            map.Add("bigint", 0);
            map.Add("float", 0);
            map.Add("decimal", 0);
            map.Add("numeric", 0);
            map.Add("varchar", null);
            map.Add("nvarchar", null);
            map.Add("char", null);
            map.Add("nchar", null);
            map.Add("text", null);
            map.Add("ntext", null);
            return map;
        }

        int? StyleForDataType(string dataTypeName)
        {
            int? style;
            if (TypeStyleMap.TryGetValue(dataTypeName.ToLower(), out style))
                return style;
            Trace.WriteLine(string.Format("Warning: No style defined for data type: {0}", dataTypeName.ToLower()));
            return null;
        }

        class StringMap : IEnumerable<string>
        {
            Dictionary<string, int> map = new Dictionary<string, int>();
            List<string> list = new List<string>();

            public int Store(string str)
            {
                int id;
                if (map.TryGetValue(str, out id))
                    return id;
                map.Add(str, id = list.Count);
                list.Add(str);
                return id++;
            }

            public int Count
            {
                get
                {
                    return list.Count;
                }
            }

            public IEnumerator<string> GetEnumerator()
            {
                return list.GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        class ConnectionInfo
        {
            public string Name { get; set; }
            public string Caption { get; set; }
            public string Database { get; set; }
            public string Server { get; set; }

            public string ConnectionString
            {
                get
                {
                    return string.Format(@"Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog={2};Data Source={1};Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False", Name, Server, Database);
                }
            }

            public string DotNetConnectionString
            {
                get
                {
                    return string.Format(@"Data Source={1}; Initial Catalog={2}; User Id=sa; Password=D1visi0ns; Application Name=WebServices;", Name, Server, Database);
                }
            }

            public override string ToString()
            {
                return Name;
            }
        }

        public MainForm()
        {
            // TODO: initialize form with contents of clipboard
            InitializeComponent();
            connectionDropdown.Items.AddRange(Connections);
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string name = "QueryResults";
                string extension = ".xlsx";
                string filename = Path.Combine(docs, name + extension);

                ConnectionInfo ci = (ConnectionInfo)connectionDropdown.SelectedItem;
                using (SqlConnection conn = new SqlConnection(ci.DotNetConnectionString))
                {
                    conn.Open();

                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = queryTextBox.Text;
                        cmd.CommandTimeout = 1000 * 60 * 10;

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            int columns = reader.FieldCount;

                            MemoryStream queryTableFile = new MemoryStream();
                            MemoryStream tableFile = new MemoryStream();
                            MemoryStream sheetFile = new MemoryStream();
                            MemoryStream sharedStringsFile = new MemoryStream();
                            MemoryStream workbookFile = new MemoryStream();

                            StringMap strMap = new StringMap();

                            int rowCount = BuildSheetFile(reader, strMap, sheetFile);
                            BuildSharedStringsFile(strMap, sharedStringsFile);
                            BuildQueryTableFile(reader, queryTableFile);
                            BuildTableFile(reader, rowCount, tableFile);
                            BuildWorkbookFile(reader, rowCount, workbookFile);

                            // basically zip up everything in the template directory, doing template replacement in appropriate places
                            using (ZipFile zip = new ZipFile())
                            {
                                // TODO: fix this, load from resource or something
                                Add(zip, TemplateFolder, "", delegate(string sourcePath, string targetFolder)
                                {
                                    Trace.WriteLine(sourcePath + "   -->    " + targetFolder);
                                    if (targetFolder == @"xl/queryTables" && Path.GetFileName(sourcePath) == "queryTable1.xml")
                                    {
                                        return queryTableFile;
                                    }
                                    else if (targetFolder == @"xl/tables" && Path.GetFileName(sourcePath) == "table1.xml")
                                    {
                                        return tableFile;
                                    }
                                    else if (targetFolder == @"xl" && Path.GetFileName(sourcePath) == "sharedStrings.xml")
                                    {
                                        return sharedStringsFile;
                                    }
                                    else if (targetFolder == @"xl" && Path.GetFileName(sourcePath) == "workbook.xml")
                                    {
                                        return workbookFile;
                                    }
                                    else if (targetFolder == @"xl/worksheets" && Path.GetFileName(sourcePath) == "sheet1.xml")
                                    {
                                        return sheetFile;
                                    }
                                    else if (targetFolder == "xl" && Path.GetFileName(sourcePath) == "connections.xml")
                                    {
                                        XmlDocument doc = new XmlDocument();
                                        doc.Load(sourcePath);
                                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                                        nsmgr.AddNamespace("x", doc.DocumentElement.GetAttribute("xmlns"));

                                        /*
                                         * now setting the following to "Query"
                                        XmlElement c = (XmlElement)doc.SelectSingleNode("/x:connections/x:connection", nsmgr);
                                        c.SetAttribute("name", ci.Name);
                                        c.SetAttribute("description", ci.Name);
                                         */

                                        XmlElement dbPr = (XmlElement)doc.SelectSingleNode("/x:connections/x:connection/x:dbPr", nsmgr);
                                        dbPr.SetAttribute("command", CommandTextToAttribute(queryTextBox.Text));
                                        dbPr.SetAttribute("connection", ci.ConnectionString);

                                        MemoryStream file = new MemoryStream();
                                        doc.Save(file);
                                        file.Position = 0;
                                        return file;
                                    }
                                    return null;
                                });
                                // TODO: fix this, save to different filename if first exists already
                                int num = 0;
                                while (File.Exists(filename))
                                {
                                    num++;
                                    filename = Path.Combine(docs, name + num.ToString() + extension);
                                }

                                zip.Save(filename);
                            }
                        }
                    }

                }

                ProcessStartInfo psi = new ProcessStartInfo(filename);
                psi.UseShellExecute = true;
                psi.Verb = "Open";
                Process p = Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BuildWorkbookFile(SqlDataReader reader, int rowCount, MemoryStream workbookFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Path.Combine(TemplateFolder, "xl", "workbook.xml"));
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", doc.DocumentElement.GetAttribute("xmlns"));

            XmlElement definedName = (XmlElement)doc.SelectSingleNode("/x:workbook/x:definedNames/x:definedName", nsmgr);
            definedName.InnerText = string.Format("Sheet1!$A$1:${0}${1}", GetColumnLetter(reader.FieldCount), rowCount + 1);

            XmlTextWriter wr = new XmlTextWriter(workbookFile, Encoding.UTF8);
            wr.Formatting = Formatting.None;
            doc.Save(wr);
            wr.Flush();
            workbookFile.Position = 0;
        }

        private void BuildSharedStringsFile(StringMap strMap, MemoryStream sharedStringsFile)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement sst = doc.CreateElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            doc.AppendChild(sst);
            sst.SetAttribute("count", strMap.Count.ToString());
            sst.SetAttribute("uniqueCount", strMap.Count.ToString());
            foreach (var str in strMap)
            {
                XmlElement si = doc.CreateElement("si", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                sst.AppendChild(si);

                XmlElement t = doc.CreateElement("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                t.InnerText = str;
                si.AppendChild(t);
            }

            XmlTextWriter wr = new XmlTextWriter(sharedStringsFile, Encoding.UTF8);
            wr.Formatting = Formatting.None;
            doc.Save(wr);
            wr.Flush();
            sharedStringsFile.Position = 0;

        }

        private int BuildSheetFile(SqlDataReader reader, StringMap stringMap, MemoryStream sheetFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.Load(Path.Combine(TemplateFolder, "xl", "worksheets", "sheet1.xml"));
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", doc.DocumentElement.GetAttribute("xmlns"));

            float[] widths = new float[reader.FieldCount];
            XmlElement[] colElements = new XmlElement[widths.Length];

            XmlElement cols = doc.CreateElement("cols", doc.DocumentElement.GetAttribute("xmlns"));
            for (int i = 0; i < reader.FieldCount; i++)
            {
                XmlElement col = doc.CreateElement("col", doc.DocumentElement.GetAttribute("xmlns"));
                col.SetAttribute("min", (i + 1).ToString());
                col.SetAttribute("max", (i + 1).ToString());
                // TODO: maybe try to pre-calculate
                col.SetAttribute("width", "11");
                col.SetAttribute("bestFit", "1");
                col.SetAttribute("customWidth", "1");
                cols.AppendChild(col);
                colElements[i] = col;
            }
            Replace(doc, "/x:worksheet/x:cols", nsmgr, cols);

            XmlElement sheetData = doc.CreateElement("sheetData", doc.DocumentElement.GetAttribute("xmlns"));
            Replace(doc, "/x:worksheet/x:sheetData", nsmgr, sheetData);

            Graphics g = CreateGraphics();
            FontFamily ff = new FontFamily("Calibri");
            Font font = new Font(ff, 11f);

            // emit header row
            {
                XmlElement row = doc.CreateElement("row", doc.DocumentElement.GetAttribute("xmlns"));
                row.SetAttribute("r", "1");
                row.SetAttribute("spans", string.Format("1:{0}", reader.FieldCount));
                row.SetAttribute("dyDescent", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "0.25");
                sheetData.AppendChild(row);

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    XmlElement c = doc.CreateElement("c", doc.DocumentElement.GetAttribute("xmlns"));
                    c.SetAttribute("r", GetColumnLetter(i + 1) + "1");
                    c.SetAttribute("t", "s");

                    XmlElement v = doc.CreateElement("v", doc.DocumentElement.GetAttribute("xmlns"));
                    string cellValue = reader.GetName(i);
                    if (cellValue == "")
                        cellValue = string.Format("Column{0}", i + 1);
                    if (cellValue != "")
                    {
                        widths[i] = Math.Max(widths[i], g.MeasureString(cellValue, font).Width + 18f); // 18f to include filter dropdown button
                        int id = stringMap.Store(cellValue);
                        v.InnerText = id.ToString();
                        row.AppendChild(c);
                        c.AppendChild(v);
                    }
                }
            }

            int rowCount = 0;
            while (reader.Read())
            {
                int rowNum = rowCount + 2;
                XmlElement row = doc.CreateElement("row", doc.DocumentElement.GetAttribute("xmlns"));
                row.SetAttribute("r", rowNum.ToString());
                row.SetAttribute("spans", string.Format("1:{0}", reader.FieldCount));
                row.SetAttribute("dyDescent", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "0.25");
                sheetData.AppendChild(row);

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    object value = reader.GetValue(i);
                    int? style = StyleForDataType(reader.GetDataTypeName(i));
                    XmlElement c = doc.CreateElement("c", doc.DocumentElement.GetAttribute("xmlns"));
                    c.SetAttribute("r", GetColumnLetter(i + 1) + rowNum.ToString());
                    c.SetAttribute("t", "s");
                    if (style != null)
                        c.SetAttribute("s", (style + 1).ToString());

                    XmlElement v = doc.CreateElement("v", doc.DocumentElement.GetAttribute("xmlns"));
                    string cellValue = GetStringValue(reader, i);
                    if (cellValue != "")
                    {
                        widths[i] = Math.Max(widths[i], g.MeasureString(cellValue, font).Width);
                        int id = stringMap.Store(cellValue);
                        v.InnerText = id.ToString();
                        row.AppendChild(c);
                        c.AppendChild(v);
                    }
                }

                rowCount++;
            }


            // set the column widths
            for (int i = 0; i < widths.Length; i++)
            {
                colElements[i].SetAttribute("width", (widths[i] / 7).ToString());
            }

            XmlElement dimension = (XmlElement)doc.SelectSingleNode("/x:worksheet/x:dimension", nsmgr);
            dimension.SetAttribute("ref", string.Format("A1:{0}{1}", GetColumnLetter(reader.FieldCount), rowCount + 1));

            XmlTextWriter wr = new XmlTextWriter(sheetFile, new UTF8Encoding(false));
            wr.Formatting = Formatting.None;
            doc.Save(wr);
            wr.Flush();
            sheetFile.Position = 0;

            return rowCount;
        }

        private void Replace(XmlDocument doc, string path, XmlNamespaceManager nsmgr, XmlElement replacement)
        {
            XmlNode node = doc.SelectSingleNode(path, nsmgr);
            node.ParentNode.ReplaceChild(replacement, node);
        }

        private string GetStringValue(SqlDataReader reader, int i)
        {
            if (reader.IsDBNull(i))
                return String.Empty;
            object obj = reader.GetValue(i);
            return obj.ToString();
        }

        private void BuildTableFile(SqlDataReader reader, int rowCount, MemoryStream tableFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Path.Combine(TemplateFolder, "xl", "tables", "table1.xml"));
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", doc.DocumentElement.GetAttribute("xmlns"));

            XmlElement table = doc.DocumentElement;
            // TODO: get number of rows
            table.SetAttribute("ref", string.Format("A1:{0}{1}", GetColumnLetter(reader.FieldCount), rowCount + 1));

            // TODO: get number of rows
            XmlElement autoFilter = (XmlElement)table.SelectSingleNode("x:autoFilter", nsmgr);
            autoFilter.SetAttribute("ref", string.Format("A1:{0}{1}", GetColumnLetter(reader.FieldCount), rowCount + 1));

            XmlElement tableColumns = (XmlElement)table.SelectSingleNode("x:tableColumns", nsmgr);
            tableColumns.SetAttribute("count", string.Format("{0}", reader.FieldCount));

            // remove all existing fields from template
            while (tableColumns.FirstChild != null)
                tableColumns.RemoveChild(tableColumns.FirstChild);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                string name = reader.GetName(i);
                int? style = StyleForDataType(reader.GetDataTypeName(i));
                if (name == "")
                    name = string.Format("Column{0}", i + 1);
                XmlElement tableColumn = doc.CreateElement("tableColumn", doc.DocumentElement.GetAttribute("xmlns"));
                tableColumn.SetAttribute("id", string.Format("{0}", i + 1));
                tableColumn.SetAttribute("uniqueName", string.Format("{0}", i + 1));
                tableColumn.SetAttribute("name", name);
                tableColumn.SetAttribute("queryTableFieldId", string.Format("{0}", i + 1));
                if (style != null)
                    tableColumn.SetAttribute("dataDxfId", style.ToString());
                tableColumns.AppendChild(tableColumn);
            }

            XmlTextWriter wr = new XmlTextWriter(tableFile, Encoding.UTF8);
            wr.Formatting = Formatting.None;
            doc.Save(wr);
            wr.Flush();
            tableFile.Position = 0;
        }

        private void BuildQueryTableFile(SqlDataReader reader, MemoryStream queryTableFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Path.Combine(TemplateFolder, "xl", "queryTables", "queryTable1.xml"));
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", doc.DocumentElement.GetAttribute("xmlns"));

            XmlElement queryTableRefresh = (XmlElement)doc.SelectSingleNode("/x:queryTable/x:queryTableRefresh", nsmgr);
            queryTableRefresh.SetAttribute("nextId", string.Format("{0}", reader.FieldCount + 1));

            XmlElement queryTableFields = (XmlElement)queryTableRefresh.SelectSingleNode("x:queryTableFields", nsmgr);
            queryTableFields.SetAttribute("count", string.Format("{0}", reader.FieldCount));

            // remove all existing fields from template
            while (queryTableFields.FirstChild != null)
                queryTableFields.RemoveChild(queryTableFields.FirstChild);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                XmlElement queryTableField = doc.CreateElement("queryTableField", doc.DocumentElement.GetAttribute("xmlns"));
                queryTableField.SetAttribute("id", string.Format("{0}", i + 1));
                queryTableField.SetAttribute("name", reader.GetName(i));
                queryTableField.SetAttribute("tableColumnId", string.Format("{0}", i + 1));
                queryTableFields.AppendChild(queryTableField);
            }

            XmlTextWriter wr = new XmlTextWriter(queryTableFile, Encoding.UTF8);
            wr.Formatting = Formatting.None;
            doc.Save(wr);
            wr.Flush();
            queryTableFile.Position = 0;
        }

        private string GetColumnLetter(int p)
        {
            int l = (p - 1) % 26;
            int h = (p - 1 - l) / 26;
            if (h == 0)
                return string.Format("{0}", (char)('A' + l));
            else
                return string.Format("{1}{0}", (char)('A' + l), (char)('A' + h));
        }

        private readonly static Regex CommandTextParse = new Regex(@"[\r\n]", RegexOptions.Compiled | RegexOptions.Multiline);
        private string CommandTextToAttribute(string sql)
        {
            return CommandTextParse.Replace(sql, delegate(Match match)
            {
                string replacement = String.Format("_x{0:x4}_", (int)match.Value[0]);
                return replacement;
            });
        }

        private void Add(ZipFile zip, string sourcePath, string targetFolder, Func<string, string, Stream> hook)
        {
            if (Directory.Exists(sourcePath))
            {
                foreach (string file in Directory.GetFiles(sourcePath))
                    Add(zip, file, targetFolder, hook);
                foreach (string file in Directory.GetDirectories(sourcePath))
                    Add(zip, file, Path.Combine(targetFolder, Path.GetFileName(file)).Replace(Path.DirectorySeparatorChar, '/'), hook);
                return;
            }
            string target = Path.Combine(targetFolder, Path.GetFileName(sourcePath)).Replace(Path.DirectorySeparatorChar, '/');
            Stream stream;
            stream = hook(sourcePath, targetFolder);
            try
            {
                if (stream == null)
                    stream = File.OpenRead(sourcePath);
                zip.AddEntry(target, stream);
            }
            finally
            {
                //if (stream != null)
                    //stream.Dispose();
            }
        }

        private void queryTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.A | Keys.Control))
            {
                e.SuppressKeyPress = true;
                Trace.WriteLine("Select All");
                queryTextBox.SelectAll();
                return;
            }
            if (e.KeyData == (Keys.Back | Keys.Control))
            {
                e.SuppressKeyPress = true;
                int selStart = queryTextBox.SelectionStart;
                while (selStart > 0 && queryTextBox.Text.Substring(selStart - 1, 1) == " ")
                {
                    selStart--;
                }
                int prevSpacePos = -1;
                if (selStart != 0)
                {
                    prevSpacePos = queryTextBox.Text.LastIndexOf(' ', selStart - 1);
                }
                queryTextBox.Select(prevSpacePos + 1, queryTextBox.SelectionStart - prevSpacePos - 1);
                queryTextBox.SelectedText = "";
                return;
            }

            Trace.WriteLine(string.Format("e.KeyData == (Keys.{0})", e.KeyData.ToString().Replace(", ", " | Keys.")));

        }
    }
}
