using System;
using System.Data;
using System.Windows.Forms;
using ExcelDictionary.Scripts;

namespace Excel_Reader
{
    public partial class Form1 : Form
    {
        private DictionaryDataManager dictionaryDataManager;

        public Form1()
        {
            InitializeComponent();

            SetPathToExelFileByUserChoice();

            ChangeDataGridViewApperanceOnStart();

            dictionaryDataManager.InitializeDictionaryFileData();

            ChangeDataGridViewComponentDataSource(dataGridView1, dictionaryDataManager.DT);
        }

        private void SetPathToExelFileByUserChoice()
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.xlsx*";
            choofdlog.FilterIndex = 1;
            string sFileName = "";

            while (true)
            {
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    if (sFileName != "")
                    {
                        dictionaryDataManager = new DictionaryDataManager(sFileName);
                        break;
                    }
                }

                if (sFileName == "")
                {
                    DialogResult res = MessageBoxManager.ShowRestartErrorWithResult("An exel file was not selected!!!");
                    if (res == DialogResult.Cancel)
                    {
                        ApplicationManager.CloseApplication();
                    }
                }
            }
        }

        private void ChangeDataGridViewApperanceOnStart()
        {
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        public void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                if (!PhraseFieldIsEmpty())
                    ChangeDataGridViewComponentDataSource(dataGridView1, dictionaryDataManager.SearchFor("Ang", KodTextBox.Text));
            }
        }

        private bool PhraseFieldIsEmpty()
        {
            return KodTextBox.Text == "";
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = dictionaryDataManager.SearchFor("Ang", KodTextBox.Text);
        }

        private void OnAlphabetButtonClick(object sender, EventArgs e)
        {
            Button clickedButton;

            try
            {
                clickedButton = (Button)sender;
            }
            catch(Exception)
            {
                return;
            }

            if(clickedButton.Name.Length == 7)
            {
                char searchingLetter = clickedButton.Name[clickedButton.Name.Length-1];
                ChangeDataGridViewComponentDataSource(dataGridView1, dictionaryDataManager.SearchFor("Ang", searchingLetter));
            }
            else if(clickedButton.Name.Length == 9)
            {
                ChangeDataGridViewComponentDataSource(dataGridView1, dictionaryDataManager.DT);
            }
        }

        private void ChangeDataGridViewComponentDataSource(DataGridView dataGridView, DataTable data)
        {
            dataGridView.DataSource = data;
        }
    }
}