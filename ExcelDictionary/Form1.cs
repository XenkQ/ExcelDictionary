using System;
using System.Data;
using System.Windows.Forms;
using ExcelDictionary.Scripts;

namespace Excel_Reader
{
    public partial class Form1 : Form
    {
        //TODO: ======================================================================================================
                //Można zmienić by za każdym razem nie przeładowywało danych jak są dobre.
                //Można jeszcze skrócić tą klasę (pomyśleć nad przeniesieniem Elementów).
                //Można by było zrobić wybieraną lokalizację pliku.
                //Na skończenie programu wymazać dane z dictionary.
        //TODO: ======================================================================================================

        private readonly DictionaryDataManager dictionaryDataManager = new DictionaryDataManager(@"..\..\Files\dictionary.xlsx");

        public Form1()
        {
            InitializeComponent();

            ChangeDataGridViewApperanceOnStart();

            dictionaryDataManager.InitializeDictionaryFileData();

            ChangeDataGridViewComponentDataSource(dataGridView1, dictionaryDataManager.DT);
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

                //Change later
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