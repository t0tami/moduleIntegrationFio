using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace moduleIntegrationFio
{
    public class WordWritter
    {
        public void WriteToWordTable(string filePath, int tableIndex, string lastName, string invalidChars, string validationResult)
        {
            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                wordDoc = wordApp.Documents.Open(filePath);
                Word.Table table = wordDoc.Tables[tableIndex];

                // Находим первую пустую строку
                int targetRow = FindFirstEmptyRow(table);

                // Добавляем новую строку если нужно
                if (targetRow > table.Rows.Count)
                {
                    table.Rows.Add();
                }

                // Заполняем все три столбца
                table.Cell(targetRow, 1).Range.Text = lastName;          // 1-й столбец - фамилия
                table.Cell(targetRow, 2).Range.Text = invalidChars;      // 2-й столбец - лишние символы
                table.Cell(targetRow, 3).Range.Text = validationResult;  // 3-й столбец - результат

                wordDoc.Save();
                MessageBox.Show("Данные успешно записаны в таблицу!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при записи в Word: {ex.Message}");
                throw;
            }
            finally
            {
                if (wordDoc != null) wordDoc.Close();
                if (wordApp != null) wordApp.Quit();
            }
        }

        private int FindFirstEmptyRow(Word.Table table)
        {
            // Ищем первую строку, где все три столбца пусты
            for (int row = 3; row <= table.Rows.Count; row++)
            {
                if (string.IsNullOrWhiteSpace(table.Cell(row, 1).Range.Text) &&
                    string.IsNullOrWhiteSpace(table.Cell(row, 2).Range.Text) &&
                    string.IsNullOrWhiteSpace(table.Cell(row, 3).Range.Text))
                {
                    return row;
                }
            }
            return table.Rows.Count + 1;
        }
    }
}
