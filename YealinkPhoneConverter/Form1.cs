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

            for (int i = 0; i < fileLines.Length; i++)
            {
                var line = fileLines[i];

                if (line.Contains("</DirectoryEntry>"))
                {
                    Contact contact = new Contact();
                    contact.FirstName = fileLines[i - 2].Replace("<Name>", "").Replace("</Name>", "").Trim();
                    contact.Work = fileLines[i - 1].Replace("<Telephone>", "").Replace("</Telephone>", "").Trim();

                    bindingList.Add(contact);
                }
            }
        }

        private void ParseCSV(string path)
        {
            string[] fileLines = File.ReadAllLines(path);

            string[] headerNames = fileLines[0].Split(',');

            for (int i = 0; i < headerNames.Length; i++)
            {
                headerNames[i] = headerNames[i].Trim();
            }

            for (int i = 1; i < fileLines.Length; i++)
            {
                string[] columns = fileLines[i].Split(',');

                Contact contact = new Contact();
                

                for (int j = 0; j < headerNames.Length; j++)
                {
                    switch (headerNames[j])
                    {
                        case "FirstName":
                            contact.FirstName = columns[j].Trim();
                            break;

                        case "LastName":
                            contact.LastName = columns[j].Trim();
                            break;

                        case "Mobile":
                            contact.Mobile = columns[j].Trim();
                            break;

                        case "Work":
                            contact.Work = columns[j].Trim();
                            break;
                    }
                }

                MessageBox.Show($"Added contact:\nName: {contact.FirstName} {contact.LastName}\nMobile: {contact.Mobile}\nWork: {contact.Work}");

                bindingList.Add(contact);
            }

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

                if (contact.Mobile != "")
                    sb.AppendLine($"\t\t<Telephone>{contact.Mobile}</Telephone>");
                else
                    sb.AppendLine($"\t\t<Telephone>{contact.Work}</Telephone>");

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


    }
}
