using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Linq;
using System.Runtime.Serialization;

namespace YealinkPhoneConverter
{
    public partial class Form1 : Form
    {
        string fileName;

        BindingList<Contact> bindingList = new BindingList<Contact>();
        


        public Form1()
        {
            InitializeComponent();

            BindingSource source = new BindingSource(bindingList, null);
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = source;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Displays an OpenFileDialog so the user can select a Cursor.  
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "CSV Files|*.CSV";
            openFileDialog1.Title = "Select a Phonebook File";

            // Show the Dialog.  
            // If the user clicked OK in the dialog and  
            // a .CUR file was selected, open it.  
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.InitialDirectory + openFileDialog1.FileName;
                ParseCSV(fileName);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportXML();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Displays an OpenFileDialog so the user can select a Cursor.  
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "XML Files|*.XML";
            openFileDialog1.Title = "Select a Phonebook File";

            // Show the Dialog.  
            // If the user clicked OK in the dialog and  
            // a .CUR file was selected, open it.  
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.InitialDirectory + openFileDialog1.FileName;
                ParseXML(fileName);
            }
        }

        private void ParseXML(string path)
        {
            string[] fileLines = File.ReadAllLines(path);
            int addedCount = 0;

            for (int i = 0; i < fileLines.Length; i++)
            {
                var line = fileLines[i];

                if (line.Contains("</DirectoryEntry>"))
                {
                    Contact contact = new Contact();
                    contact.FirstName = RemoveSpecialCharacters(fileLines[i - 2].Replace("<Name>", "").Replace("</Name>", "").Trim());
                    contact.Work = RemoveSpecialCharacters(fileLines[i - 1].Replace("<Telephone>", "").Replace("</Telephone>", "").Trim());

                    bindingList.Add(contact);
                    addedCount++;
                }
            }

            MessageBox.Show($"Added {addedCount} contacts.");
        }

        private void ParseCSV(string path)
        {
            string[] fileLines = File.ReadAllLines(path);

            string[] headerNames = fileLines[0].Split(',');

            for (int i = 0; i < headerNames.Length; i++)
            {
                headerNames[i] = headerNames[i].Trim();
            }

            int addedCount = 0;

            for (int i = 1; i < fileLines.Length; i++)
            {
                string[] columns = fileLines[i].Split(',');

                Contact contact = new Contact();
                

                for (int j = 0; j < headerNames.Length; j++)
                {
                    switch (headerNames[j])
                    {
                        case "FirstName":
                            contact.FirstName = RemoveSpecialCharacters(columns[j].Trim());
                            break;

                        case "LastName":
                            contact.LastName = RemoveSpecialCharacters(columns[j].Trim());
                            break;

                        case "Mobile":
                            contact.Mobile = RemoveSpecialCharacters(columns[j].Trim());
                            break;

                        case "Work":
                            contact.Work = RemoveSpecialCharacters(columns[j].Trim());
                            break;
                    }
                }

                bindingList.Add(contact);
                addedCount++;
            }

            MessageBox.Show($"Added {addedCount} contacts.");
        }

        private void ExportXML()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<CiscoIPPhoneDirectory>");
            sb.AppendLine("\t<Title>Phonebook</Title>");
            sb.AppendLine("\t<Prompt>Phonebook</Prompt>");


            foreach (Contact contact in bindingList)
            {
                sb.AppendLine("\t<DirectoryEntry>");
                sb.AppendLine($"\t\t<Name>{contact.FirstName} {contact.LastName}</Name>");

                if (!string.IsNullOrWhiteSpace(contact.Work))
                    sb.AppendLine($"\t\t<Telephone>{contact.Work}</Telephone>");
                else
                    sb.AppendLine($"\t\t<Telephone>{contact.Mobile}</Telephone>");

                sb.AppendLine("\t</DirectoryEntry>");
            }
            sb.AppendLine("</CiscoIPPhoneDirectory>");

            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "XML Files|*.XML";
            fileDialog.Title = "Save phonebook";
            fileDialog.FileName = "Phonebook.XML";
            
            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (fileDialog.FileName != "")
                {
                    FileStream file = File.Create(fileDialog.InitialDirectory + fileDialog.FileName);

                    byte[] info = new UTF8Encoding(true).GetBytes(sb.ToString());
                    file.Write(info, 0, info.Length);
                    file.Close();
                }
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (var a in bindingList
                    .ToLookup(x => new { x.FirstName, x.LastName, x.Mobile, x.Work })
                    .SelectMany(x => x.Skip(1))
                    .ToArray())
            {
                bindingList.Remove(a);
            }
        }

        private static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < str.Length; i++)
            {
                if ((str[i] >= '0' && str[i] <= '9')
                    || (str[i] >= 'A' && str[i] <= 'z'
                        || (str[i] == '.' || str[i] == '_'
                            || (str[i] == '+' || str[i] == '\''))))
                {
                    sb.Append(str[i]);
                }
            }

            return sb.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            bindingList.Clear();
        }
    }
}
