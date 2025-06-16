using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace moduleIntegrationFio
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ServerRequest request = new ServerRequest();
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void GetRequestButtonClick(object sender, RoutedEventArgs e)
        {
            string url = "http://localhost:4444/TransferSimulator/fullName";

            string result = await request.GetRequestAsync(url);

            DataTextBlock.Text = GetEmail(result);
        }

        private string GetEmail(string result)
        {
            return result.Substring(result.IndexOf(":") + 2)
                .Replace("\"", "")
                .Replace("}", "");
        }

        private void SendResultButtonClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(DataTextBlock.Text))
            {
                MessageBox.Show("Данные с сервера ещё не получены!");
                return;
            }

            string filePath = @"C:\Users\Edemskaya.as\source\repos\moduleIntegrationFio\ТестКейс.docx";
            string lastName = DataTextBlock.Text.Trim();

            // Проверка валидности фамилии
            bool isValid = Regex.IsMatch(lastName, @"^[a-zA-Zа-яА-ЯёЁ\s]+$");
            string validationResult = isValid ? "успешно" : "не успешно";

            try
            {
                WordWritter wordWritter = new WordWritter();
                wordWritter.WriteToWordTable(filePath, 1, lastName, validationResult);

                ResultTextBlock.Text = isValid
                    ? "Данные успешно записаны в таблицу!"
                    : "Данные записаны с пометкой 'не успешно'!";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
    }
}
