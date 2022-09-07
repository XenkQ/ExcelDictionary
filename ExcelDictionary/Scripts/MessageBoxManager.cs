using System.Windows.Forms;

namespace ExcelDictionary.Scripts
{
    public static class MessageBoxManager
    {
        public static void ShowErrorMessageBox(string message)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void ShowErrorMessageBoxWithChoice(string message, string choice1, string choice2)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void ShowExclamationMessageBox(string message)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}
