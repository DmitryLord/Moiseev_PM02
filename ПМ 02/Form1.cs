using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace ПМ_02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OpenFileDialog ofd = new OpenFileDialog();

        double money = 0;
        public void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
            {
                MessageBox.Show("Ошибка");
            }
            else
            {
                if (radioButton1.Checked)
                {
                    money = money + ((money * 50) / 100);
                    label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                }
                else if (radioButton2.Checked)
                {
                    label3.Text = $"{money + ((money * 7) / 100)}";
                    label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                }
                else
                {
                    label3.Text = $"{money + ((money * 20) / 100)}";
                    label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                }
                label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                button2.Enabled = true;
                if (numericUpDown1.Value > 10)
                {
                    label3.Text = $"{money - ((money * 5) / 100)}";
                    label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                }
                else if (numericUpDown1.Value > 15)
                {
                    label3.Text = $"{money - ((money * 7) / 100)}";
                    label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                }
                else if (numericUpDown1.Value > 20)
                {
                    label3.Text = $"{money - ((money * 10) / 100)}";
                    label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                }
                else if (numericUpDown1.Value > 30)
                {
                    label3.Text = $"{money - ((money * 25) / 100)}";
                    label3.Text = $"{money * Convert.ToDouble(numericUpDown1.Value)} руб.";
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    money = 750;
                    break;

                    case 1:
                    money = 800;
                    break;

                    case 2:
                    money = 870;
                    break;
                case 3:
                    money = 950;
                    break;
                case 4:
                    money = 600;
                    break;
            }
                
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Создаём объект документа
            Word.Document doc = null;
            try
            {
                // Создаём объект приложения
                Word.Application app = new Word.Application();
                // Путь до шаблона документа
                string source = Path.Combine(Directory.GetCurrentDirectory(), "Шаблон.docx");
                // Открываем
                doc = app.Documents.Add(source);
                doc.Activate();

                // Добавляем информацию
                // wBookmarks содержит все закладки
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                Word.Range wRange;
                int i = 0;
                int num = 0;
                num++;
                string nm = Convert.ToString(num);
                string[] data = new string[4] {DateTime.Now.ToShortDateString(), label3.Text, nm ,comboBox1.Items[comboBox1.SelectedIndex].ToString()};
                foreach (Word.Bookmark mark in wBookmarks)
                {

                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }

                // Закрываем документ
                MessageBox.Show("Квитанция успешно оформлена!", "Квитанция", MessageBoxButtons.OK, MessageBoxIcon.Information);
                doc.Close();
                doc = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Image Files(*.JPG;*.JPEG;)|*.JPG;*.JPEG; | All files(*.*) | *.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    pictureBox1.Image = new Bitmap(ofd.FileName);
                }
                catch
                {
                    MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка");
                }
            }
        }
    }
}
