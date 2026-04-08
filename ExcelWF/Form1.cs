using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelWF
{

    public partial class Form1 : Form
    {
        // Путь к текущему Excel файлу
        private string currentExcelFilePath;
        
        // Имя листа в Excel
        private const string ExcelSheetName = "Data";

        public Form1()
        {
            InitializeComponent();
            
            // Устанавливаем лицензионный контекст для EPPlus (для некоммерческого использования)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            // Инициализация начального состояния
            currentExcelFilePath = string.Empty;
            UpdateStatus("Готов к работе. Создайте новый Excel файл или откройте существующий.");
        }

        private void btnCreateExcel_Click(object sender, EventArgs e)
        {
            try
            {
                // Настройка диалога сохранения файла
                saveFileDialog.FileName = $"PersonData_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.Title = "Создать новый Excel файл";

                // Показ диалога сохранения
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    currentExcelFilePath = saveFileDialog.FileName;
                    
                    // Создание нового Excel файла
                    CreateNewExcelFile(currentExcelFilePath);
                    
                    UpdateStatus($"Файл создан: {currentExcelFilePath}");
                    MessageBox.Show("Excel файл успешно создан!", "Успех", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                HandleError("Ошибка при создании файла", ex);
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Заполнить данные"
        /// </summary>
        private void btnFillData_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверка валидности данных
                if (!ValidateInputData())
                {
                    return;
                }

                // Если файл не выбран, предлагаем создать новый
                if (string.IsNullOrEmpty(currentExcelFilePath))
                {
                    var result = MessageBox.Show(
                        "Excel файл не выбран. Хотите создать новый файл?",
                        "Выберите файл",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        btnCreateExcel_Click(sender, e);
                        if (string.IsNullOrEmpty(currentExcelFilePath))
                        {
                            return; // Пользователь отменил создание файла
                        }
                    }
                    else
                    {
                        return;
                    }
                }

                // Заполнение данных в Excel
                FillDataInExcel(currentExcelFilePath);
                
                UpdateStatus($"Данные успешно записаны в файл: {currentExcelFilePath}");
                MessageBox.Show("Данные успешно записаны в Excel файл!", "Успех", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                HandleError("Ошибка при заполнении данных", ex);
            }
        }

        /// <summary>
        /// Создание нового Excel файла с заголовками
        /// </summary>
        private void CreateNewExcelFile(string filePath)
        {
            // Удаляем существующий файл если он есть
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            // Создаем новый пакет Excel
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Добавление листа
                var worksheet = package.Workbook.Worksheets.Add(ExcelSheetName);

                // Настройка ширины столбцов
                worksheet.Column(1).Width = 25;  // Имя
                worksheet.Column(2).Width = 25;  // Фамилия
                worksheet.Column(3).Width = 20;  // Дата рождения
                worksheet.Column(4).Width = 30;  // Email
                worksheet.Column(5).Width = 20;  // Телефон
                worksheet.Column(6).Width = 40;  // Адрес
                worksheet.Column(7).Width = 25;  // Должность
                worksheet.Column(8).Width = 25;  // Отдел
                worksheet.Column(9).Width = 50;  // Заметки

                // Создание заголовков с форматированием
                CreateHeaders(worksheet);

                // Сохранение файла
                package.Save();
            }
        }

        private void CreateHeaders(ExcelWorksheet worksheet)
        {
            // Массив заголовков
            string[] headers = new string[]
            {
                "Имя",
                "Фамилия",
                "Дата рождения",
                "Email",
                "Телефон",
                "Адрес",
                "Должность",
                "Отдел",
                "Заметки"
            };

            // Применение заголовков и стилей
            for (int col = 1; col <= headers.Length; col++)
            {
                var cell = worksheet.Cells[1, col];
                cell.Value = headers[col - 1];
                
                // Жирный шрифт для заголовков
                cell.Style.Font.Bold = true;
                cell.Style.Font.Size = 12;
                
                // Заливка заголовка
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                
                // Границы
                cell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                cell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                cell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                cell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                
                // Выравнивание по центру
                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
        }


        private void FillDataInExcel(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Файл не найден: {filePath}");
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Получение листа
                var worksheet = package.Workbook.Worksheets[ExcelSheetName];
                
                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Лист с именем '{ExcelSheetName}' не найден в файле");
                }

                // Поиск последней заполненной строки
                int lastRow = worksheet.Dimension?.End.Row ?? 0;
                int newRow = lastRow + 1;

                // Заполнение данных в новой строке
                worksheet.Cells[newRow, 1].Value = txtFirstName.Text.Trim();
                worksheet.Cells[newRow, 2].Value = txtLastName.Text.Trim();
                worksheet.Cells[newRow, 3].Value = dtpBirthDate.Value.ToString("dd.MM.yyyy");
                worksheet.Cells[newRow, 4].Value = txtEmail.Text.Trim();
                worksheet.Cells[newRow, 5].Value = txtPhone.Text.Trim();
                worksheet.Cells[newRow, 6].Value = txtAddress.Text.Trim();
                worksheet.Cells[newRow, 7].Value = txtPosition.Text.Trim();
                worksheet.Cells[newRow, 8].Value = txtDepartment.Text.Trim();
                worksheet.Cells[newRow, 9].Value = txtNotes.Text.Trim();

                // Применение границ к новой строке
                for (int col = 1; col <= 9; col++)
                {
                    var cell = worksheet.Cells[newRow, col];
                    cell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    cell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    cell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    cell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }

                // Автоматическая настройка ширины столбцов
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Сохранение файла
                package.Save();
            }
        }

        private bool ValidateInputData()
        {
            // Проверка обязательных полей
            if (string.IsNullOrWhiteSpace(txtFirstName.Text))
            {
                MessageBox.Show("Пожалуйста, введите имя.", "Ошибка валидации", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtFirstName.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtLastName.Text))
            {
                MessageBox.Show("Пожалуйста, введите фамилию.", "Ошибка валидации", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtLastName.Focus();
                return false;
            }

            // Валидация Email
            if (!string.IsNullOrWhiteSpace(txtEmail.Text))
            {
                try
                {
                    var mailAddress = new System.Net.Mail.MailAddress(txtEmail.Text);
                }
                catch
                {
                    MessageBox.Show("Пожалуйста, введите корректный Email.", "Ошибка валидации", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtEmail.Focus();
                    return false;
                }
            }

            return true;
        }


        private void UpdateStatus(string message)
        {
            tsslStatus.Text = message;
        }

        private void HandleError(string context, Exception ex)
        {
            string errorMessage = $"{context}: {ex.Message}";
            UpdateStatus($"Ошибка: {ex.Message}");
            MessageBox.Show(errorMessage, "Ошибка", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}

