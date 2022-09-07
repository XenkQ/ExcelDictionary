using System.Windows.Forms;

namespace ExcelDictionary.Scripts
{
    public static class MessageBoxManager
    {
        public static void ShowErrorMessageBox(string message)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static DialogResult ShowRestartErrorWithResult(string message)
        {
            return MessageBox.Show(message, "", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
        }

        public static void ShowExclamationMessageBox(string message)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}
