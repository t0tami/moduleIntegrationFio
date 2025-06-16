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
        public void WriteToWordTable(string filePath, int tableIndex, string lastName, string validationResult)
        {
            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                wordDoc = wordApp.Documents.Open(filePath);
                Word.Table table = wordDoc.Tables[tableIndex];

                // Находим первую полностью пустую строку
                int targetRow = FindFirstEmptyRow(table);

                // Если все строки заполнены, добавляем новую
                if (targetRow > table.Rows.Count)
                {
                    table.Rows.Add();
                }

                // Записываем данные в одну строку
                table.Cell(targetRow, 1).Range.Text = lastName;      // 1-й столбец - фамилия
                table.Cell(targetRow, 3).Range.Text = validationResult; // 3-й столбец - результат

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
            // Проверяем строки начиная с 3-й (как в вашем оригинальном коде)
            for (int row = 3; row <= table.Rows.Count; row++)
            {
                // Считаем строку пустой, если оба столбца пусты
                if (string.IsNullOrWhiteSpace(table.Cell(row, 1).Range.Text) &&
                    string.IsNullOrWhiteSpace(table.Cell(row, 3).Range.Text))
                {
                    return row;
                }
            }
            return table.Rows.Count + 1; // Возвращаем следующую после последней
        }
    }
}
