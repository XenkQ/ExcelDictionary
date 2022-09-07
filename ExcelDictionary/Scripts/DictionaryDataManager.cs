using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelDictionary.Scripts
{
    public class DictionaryDataManager
    {
        private DataTableCollection dataTableCollection;
        private DataTable dt;
        public DataTable DT { get { return dt; } }
        private DataTable localData = null;
        public DataTable LocalData { get { return localData; } }
        private string ExcelFilePath;
        private string currentDataToSearch = "";
        

        public DictionaryDataManager(string ExcelFilePath)
        {
            this.ExcelFilePath = ExcelFilePath;
        }

        public void InitializeDictionaryFileData()
        {
            try
            {
                using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
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
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBoxManager.ShowErrorMessageBox("The dictionary file does not exist, create a new file named \"dictionary\" in .xlsx format in the \"Files\" folder");
                ApplicationManager.CloseApplication();
            }
            catch (IOException)
            {
                MessageBoxManager.ShowErrorMessageBox("Close the dictionary file");
                ApplicationManager.CloseApplication();
            }
        }

        public void SearchForPhrase(string phrase)
        {
            phrase = phrase.ToLower();
            localData = null;

            if (phrase != "" && dt != null)
            {
                SearchFor("Ang", phrase);
            }
            else if (phrase == "")
            {
                MessageBoxManager.ShowExclamationMessageBox("The search field is empty!!! Enter a search value");
            }
        }

        public DataTable SearchFor(string col, char letter)
        {
            if (currentDataToSearch == letter.ToString()) { return localData; }
            currentDataToSearch = letter.ToString();

            letter = char.ToLower(letter);
            try
            {
                localData = dt.AsEnumerable()
               .Where(row => row.Field<String>(col) != null ? row.Field<String>(col).ToLower()[0] == char.ToLower(letter) : throw new Exception())
               .OrderBy(row => row.Field<String>(col))
               .CopyToDataTable();
            }
            catch (InvalidOperationException)
            {
                MessageBoxManager.ShowExclamationMessageBox($"There are no words with the letter {letter} in the database");
            }
            catch (Exception)
            {

            }

            return localData;
        }

        public DataTable SearchFor(string col, string searchingText)
        {
            if (currentDataToSearch == searchingText) { return localData; }
            currentDataToSearch = searchingText;

            searchingText = searchingText.ToLower();


            try
            {
                localData = dt.AsEnumerable()
               .Where(row => row.Field<String>(col) != null ? row.Field<String>(col).ToLower().Contains(searchingText) : throw new Exception())
               .OrderBy(row => row.Field<String>(col))
               .CopyToDataTable();
            }
            catch (InvalidOperationException)
            {
                MessageBoxManager.ShowExclamationMessageBox("The search word does not exist");
            }
            catch (Exception)
            {

            }

            return localData;
        }
    }
}
