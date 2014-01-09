using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        ConnectionInfo[] Connections = new ConnectionInfo[] {
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
        };

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
            // basically zip up everything in the template directory, doing template replacement in appropriate places
            using (ZipFile zip = new ZipFile())
            {
                // TODO: fix this, load from resource or something
                Add(zip, @"C:\Users\wfalk\source\QueryToExcel\Template", "", delegate(string sourcePath, string targetFolder)
                {
                    if (targetFolder == "xl" && Path.GetFileName(sourcePath) == "connections.xml")
                    {
                        ConnectionInfo ci = (ConnectionInfo)connectionDropdown.SelectedItem;

                        XmlDocument doc = new XmlDocument();
                        doc.Load(sourcePath);
                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                        nsmgr.AddNamespace("x", doc.DocumentElement.GetAttribute("xmlns"));

                        XmlElement c = (XmlElement)doc.SelectSingleNode("/x:connection");
                        c.SetAttribute("name", ci.Name);
                        c.SetAttribute("description", ci.Name);
                        
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
                zip.Save(@"C:\Users\wfalk\Documents\MyZipFile.zip");
            }
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
                    Add(zip, file, Path.Combine(targetFolder, Path.GetFileName(file)), hook);
                return;
            }
            string target = Path.Combine(targetFolder, Path.GetFileName(sourcePath));
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
    }
}
