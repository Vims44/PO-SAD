using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word; // Работа с Word
using System.IO; // Работа с файлами

namespace cursach
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Кнопка "Рассчитать"
        private void button2_Click(object sender, EventArgs e)
        {
            List<double> data;

            // Если выбран режим генерации случайных данных
            data = GetDataFromGrid();

            if (data.Count == 0)
            {
                MessageBox.Show("Нет данных для анализа.");
                return;
            }
            else
            {
                // Получение данных из таблицы (ручной ввод/загрузка файла)
                data = GetDataFromGrid();

                if (data.Count == 0)
                {
                    MessageBox.Show("Нет данных для анализа.");
                    return;
                }
            }

            // Расчёт статистики
            StringBuilder result = new StringBuilder();
            chart1.Series[0].Points.Clear();
            chart1.ChartAreas[0].AxisX.Interval = 1; // Показ каждой подписи
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Microsoft Sans Serif", 10);
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;

            // Группируем данные (сколько раз встречается каждое значение)
            int binSize = 10; // Ширина интервала 
            int minValue = (int)Math.Floor(data.Min());
            int maxValue = (int)Math.Ceiling(data.Max());

            var bins = new Dictionary<string, int>();

            for (int i = minValue; i <= maxValue; i += binSize)
            {
                string label = $"{i}–{i + binSize - 1}";
                int count = data.Count(x => x >= i && x < i + binSize);
                bins[label] = count;
            }

            // Добавление в гистограмму
            foreach (var bin in bins)
            {
                chart1.Series[0].Points.AddXY(bin.Key, bin.Value);
            }

            // Математическое ожидание
            double sum = 0;
            for (int i = 0; i < data.Count; i++)
                sum += data[i];
            double mean = sum / data.Count;

            // Дисперсия
            double sqSum = 0;
            for (int i = 0; i < data.Count; i++)
                sqSum += Math.Pow(data[i] - mean, 2);
            double variance = sqSum / data.Count;
            double stdDev = Math.Sqrt(variance); // Среднеквадратичное отклонение

            // Математическое ожидание
            if (checkBox1.Checked)
                result.AppendLine($"Математическое ожидание: {mean:F2}");

            // Среднеквадратичное отклонение
            if (checkBox2.Checked)
                result.AppendLine($"СКО: {stdDev:F2}");

            // Медиана
            if (checkBox3.Checked)
            {
                var sorted = data.OrderBy(x => x).ToList();
                double median = sorted.Count % 2 == 0
                    ? (sorted[sorted.Count / 2 - 1] + sorted[sorted.Count / 2]) / 2.0
                    : sorted[sorted.Count / 2];
                result.AppendLine($"Медиана: {median:F2}");
            }

            // Дисперсия
            if (checkBox4.Checked)
                result.AppendLine($"Дисперсия: {variance:F2}");

            // Мода
            if (checkBox5.Checked)
            {
                var mode = data.GroupBy(x => x)
                               .OrderByDescending(g => g.Count())
                               .First().Key;
                result.AppendLine($"Мода: {mode:F2}");
            }

            // Коэффициент асимметрии
            if (checkBox6.Checked)
            {
                double skew = data.Average(x => Math.Pow(x - mean, 3)) / Math.Pow(stdDev, 3);
                result.AppendLine($"Асимметрия: {skew:F2}");
            }

            // Коэффициент эксцесса
            if (checkBox7.Checked)
            {
                double kurt = data.Average(x => Math.Pow(x - mean, 4)) / Math.Pow(stdDev, 4) - 3;
                result.AppendLine($"Эксцесс: {kurt:F2}");
            }

            // Минимум
            if (checkBox8.Checked)
                result.AppendLine($"Минимум: {data.Min():F2}");

            // Максимум
            if (checkBox9.Checked)
                result.AppendLine($"Максимум: {data.Max():F2}");

            // Размах
            if (checkBox10.Checked)
            {
                double range = data.Max() - data.Min();
                result.AppendLine($"Размах: {range:F2}");
            }

            textBox2.Text = result.ToString();
        }

        // Кнопка "Загрузить файл"
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Text Files (*.txt)|*.txt|Word Files (*.docx)|*.docx|All Files (*.*)|*.*";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    List<string> numbers = new List<string>();

                    // Чтение файла
                    if (Path.GetExtension(ofd.FileName) == ".txt")
                    {
                        var content = File.ReadAllText(ofd.FileName);
                        numbers = content.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    }
                    else if (Path.GetExtension(ofd.FileName) == ".docx")
                    {
                        var wordApp = new Microsoft.Office.Interop.Word.Application();
                        var doc = wordApp.Documents.Open(ofd.FileName, ReadOnly: true);

                        string fullText = "";
                        foreach (Microsoft.Office.Interop.Word.Paragraph p in doc.Paragraphs)
                            fullText += p.Range.Text;

                        doc.Close();
                        wordApp.Quit();

                        numbers = fullText.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    }

                    // Проверка на пустоту файла
                    if (numbers.Count == 0)
                    {
                        MessageBox.Show("Файл пуст.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Подготовка таблицы
                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();

                    for (int i = 0; i < numbers.Count; i++)
                        dataGridView1.Columns.Add($"col{i + 1}", $"X{i + 1}");
                    var row = new DataGridViewRow();
                    row.CreateCells(dataGridView1);
                    bool allValid = true;
                    for (int i = 0; i < numbers.Count; i++)
                    {
                        if (double.TryParse(numbers[i], out double val))
                            row.Cells[i].Value = val;
                        else
                            allValid = false;
                    }

                    dataGridView1.Rows.Add(row);
                    textBox1.Text = numbers.Count.ToString();
                    if (!allValid)
                        MessageBox.Show("Некоторые значения не являются числами и были пропущены.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при загрузке файла: " + ex.Message);
                }
            }
        }

        // Результат анализа
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.Multiline = true;
            textBox2.Font = new Font(textBox2.Font.FontFamily, 10);
        }

        private List<double> GetDataFromGrid()
        {
            var data = new List<double>();

            if (dataGridView1.Rows.Count > 0)
            {
                DataGridViewRow row = dataGridView1.Rows[0];
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (row.Cells[i].Value != null && double.TryParse(row.Cells[i].Value.ToString(), out double val))
                        data.Add(val);
                }
            }
            return data;
        }

        // Галочки
        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            bool check = checkBox11.Checked;

            checkBox1.Checked = check;
            checkBox2.Checked = check;
            checkBox3.Checked = check;
            checkBox4.Checked = check;
            checkBox5.Checked = check;
            checkBox6.Checked = check;
            checkBox7.Checked = check;
            checkBox8.Checked = check;
            checkBox9.Checked = check;
            checkBox10.Checked = check;
        }

        // Кнопка сброс
        private void button1_Click(object sender, EventArgs e)
        {
            // Очистка таблицы
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // Очистка полей
            textBox1.Clear(); // Количество данных
            textBox2.Clear(); // Результат
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;

            // Очистка графика
            chart1.Series[0].Points.Clear();

            // Снятие всех галочек
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
        }

        // Генерация случайных чисел после выбора ввода данных
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                if (!TryGetValidCount(out int count))
                    return;

                Random rnd = new Random();

                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                for (int i = 0; i < count; i++)
                {
                    var column = new DataGridViewTextBoxColumn();
                    column.Name = $"col{i + 1}";
                    column.HeaderText = $"X{i + 1}";
                    column.Width = 50;
                    dataGridView1.Columns.Add(column);
                }

                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView1);
                for (int i = 0; i < count; i++)
                    row.Cells[i].Value = rnd.Next(1, 101);
                dataGridView1.Rows.Add(row); // Добавляем сгенерированный ряд
            }
        }

        // Вручную
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                if (int.TryParse(textBox1.Text, out int count) && count > 0 && count <= 100)
                    {
                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();

                    // Количество колонок
                    for (int i = 0; i < count; i++)
                    {
                        var column = new DataGridViewTextBoxColumn();
                        column.Name = $"col{i + 1}";
                        column.HeaderText = $"X{i + 1}";
                        column.Width = 50; 
                        dataGridView1.Columns.Add(column);
                    }

                    // Добавление одной строки для ввода
                    dataGridView1.Rows.Add();
                    // Прокрутка
                    dataGridView1.ScrollBars = ScrollBars.Both;
                }
                else MessageBox.Show("Введите количество от 1 до 100.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Справка
        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
            "Программное обеспечение для статистического анализа данных\n\n" +
            "Разработала: Карпова Анна Сергеевна\n" +
            "Студентка 2 курса, группа 23ИСП-1\n" +
            "Орский гуманитарно-технологический институт (филиал) ОГУ\n" +
            "2025 год",
            "О программе",
            MessageBoxButtons.OK, // Кнопка ОК
            MessageBoxIcon.Information); // Иконка (синий круг с восклицательным знаком)
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button3.Visible = false; // Скрытие кнопки "Загрузить файл"
            textBox2.ReadOnly = true; // Запрет на ввод пользователем
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            button3.Visible = radioButton3.Checked; // Показ кнопки, если выбран "Загрузить файл"
        }

        // Ограничение чисел в таблице (для ввода вручную)
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.FormattedValue != null && !string.IsNullOrWhiteSpace(e.FormattedValue.ToString())) // Строка должна быть не пустой и не состоять из пробелов
            {
                if (double.TryParse(e.FormattedValue.ToString(), out double value))
                {
                    if (value < -1000 || value > 1000)
                    {
                        MessageBox.Show("Введите число от -1000 до 1000.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true; // Отмена ввода
                    }
                }
                else
                {
                    MessageBox.Show("Введите корректное число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                }
            }
        }

        // Количество данных (ввод числа юзером)
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string input = textBox1.Text;

            // Удаляем всё, кроме цифр
            string digitsOnly = new string(input.Where(char.IsDigit).ToArray());

            if (input != digitsOnly)
            {
                int selectionStart = textBox1.SelectionStart - 1;
                textBox1.Text = digitsOnly;
                textBox1.SelectionStart = Math.Max(selectionStart, 0); // Чтобы курсор не прыгал в начало
            }

            // Проверка значения
            if (int.TryParse(digitsOnly, out int value))
            {
                if (value > 100)
                {
                    MessageBox.Show("Введите число от 1 до 100", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox1.Text = "";
                    textBox1.SelectionStart = textBox1.Text.Length;
                }
            }
        }

        private void TextBox_KeyPress_OnlyNumbers(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;

            // Разрешаем цифры, запятую, минус и управляющие символы (Backspace и т.п.)
            if (!char.IsDigit(c) && c != ',' && c != '-' && !char.IsControl(c))
            {
                e.Handled = true; // Блокирока ввода остальных символов
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is TextBox tb)
            {
                // Убираем старый обработчик (если есть)
                tb.KeyPress -= TextBox_KeyPress_OnlyNumbers;

                // Добавляем новый
                tb.KeyPress += TextBox_KeyPress_OnlyNumbers;
            }
        }
        private bool TryGetValidCount(out int count)
        {
            count = 0;
            string input = textBox1.Text.Trim();

            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show("Введите количество чисел от 1 до 100.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!int.TryParse(input, out count) || count < 1 || count > 100)
            {
                MessageBox.Show("Введите число от 1 до 100.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
    }
}
