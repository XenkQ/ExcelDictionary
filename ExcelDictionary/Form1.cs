using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Reader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            StartData();
        }

        private DataTableCollection dataTableCollection;
        private DataTable dt;
        private DataTable localData = null;
        private string pathExcelFile = @"..\..\Files\dictionary.xlsx";

        private void StartData()
        {
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            try
            {
                using (var stream = File.Open(pathExcelFile, FileMode.Open, FileAccess.Read))
                {
                    try
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            dataTableCollection = result.Tables;
                        }

                        dt = dataTableCollection[0];
                        dataGridView1.DataSource = dt;
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("The dictionary file does not exist, create a new file named \"dictionary\" in .xlsx format in the \"Files\" folder", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }
            catch (IOException)
            {
                MessageBox.Show("Close the dictionary file", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            SearchPhrase();
        }

        private void SearchPhrase()
        {
            string searchFor = KodTextBox.Text.ToLower();
            localData = null;
            if (KodTextBox.Text != "" && dt != null)
            {
                SearchFor("Ang", searchFor, dt);
            }
            else if (KodTextBox.Text == "")
            {
                MessageBox.Show("The search field is empty!!! Enter a search value", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                SearchPhrase();
            }
        }

        private void OnAlphabetButtonClick(object sender, EventArgs e)
        {
            if (sender.Equals(buttonA))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'a', dt);
            }

            if (sender.Equals(buttonB))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'b', dt);
            }

            if (sender.Equals(buttonC))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'c', dt);
            }

            if (sender.Equals(buttonD))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'd', dt);
            }

            if (sender.Equals(buttonE))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'e', dt);
            }

            if (sender.Equals(buttonF))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'f', dt);
            }

            if (sender.Equals(buttonG))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'g', dt);
            }

            if (sender.Equals(buttonH))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'h', dt);
            }

            if (sender.Equals(buttonI))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'i', dt);
            }

            if (sender.Equals(buttonJ))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'j', dt);
            }

            if (sender.Equals(buttonK))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'k', dt);
            }

            if (sender.Equals(buttonL))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'l', dt);
            }

            if (sender.Equals(buttonM))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'm', dt);
            }

            if (sender.Equals(buttonN))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'n', dt);
            }

            if (sender.Equals(buttonO))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'o', dt);
            }

            if (sender.Equals(buttonP))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'p', dt);
            }

            if (sender.Equals(buttonQ))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'q', dt);
            }

            if (sender.Equals(buttonR))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'r', dt);
            }

            if (sender.Equals(buttonS))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 's', dt);
            }

            if (sender.Equals(buttonT))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 't', dt);
            }

            if (sender.Equals(buttonU))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'u', dt);
            }

            if (sender.Equals(buttonV))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'v', dt);
            }

            if (sender.Equals(buttonW))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'w', dt);
            }

            if (sender.Equals(buttonX))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'x', dt);
            }

            if (sender.Equals(buttonY))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'y', dt);
            }

            if (sender.Equals(buttonZ))
            {
                DisplayOnlyWithGivenFirstChar("Ang", 'z', dt);
            }

            if (sender.Equals(buttonAll))
            {
                dataGridView1.DataSource = dt;
            }
        }

        private void SearchFor(string col, string searchingText, DataTable dataTable)
        {
            Console.WriteLine(searchingText);
            try
            {
                localData = dataTable.AsEnumerable()
               .Where(row => row.Field<String>(col) != null ? row.Field<String>(col).ToLower().Contains(searchingText) : throw new Exception())
               .OrderBy(row => row.Field<String>(col))
               .CopyToDataTable();
                dataGridView1.DataSource = localData;
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("The search word does not exist", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception)
            {

            }
        }

        private void DisplayOnlyWithGivenFirstChar(string col, char letter, DataTable dataTable)
        {
            try
            {
                localData = dataTable.AsEnumerable()
               .Where(row => row.Field<String>(col) != null ? row.Field<String>(col).ToLower()[0] == char.ToLower(letter) : throw new Exception())
               .OrderBy(row => row.Field<String>(col))
               .CopyToDataTable();
                dataGridView1.DataSource = localData;
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show($"There are no words with the letter {letter} in the database", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
             catch (Exception)
            {

            }
        }
    }
}
