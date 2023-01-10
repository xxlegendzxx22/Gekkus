using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ZXCGHOUL993
{
    
    public partial class Form1 : Form
    {
        bool V = false;
        bool B = false;
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;

        public Form1()
        {
            InitializeComponent();
        }
        
        private void создатьТаблицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns.RemoveAt(i);
            };
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns.RemoveAt(i);
            };
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns.RemoveAt(i);
            };
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns.RemoveAt(i);
            };
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns.RemoveAt(i);
            };
            DataGridViewTextBoxColumn[] column = new DataGridViewTextBoxColumn[18];
            for (int i = 0; i < 18; i++)
            {
                column[i] = new DataGridViewTextBoxColumn(); // выделяем память для объекта
                column[i].HeaderText = "Столбец" + i;
                column[i].Name = "Header" + i;
                
            }
            this.dataGridView1.Columns.AddRange(column);
            for (int j = 0; j < 45; j++) dataGridView1.Rows.Add();
            dataGridView1.Rows[44].Cells[17].Value = "*";
            dataGridView1.Rows[0].Cells[0].Selected = false;
            dataGridView1.Rows[1].Cells[2].Selected = true;
            B = true;
        }


        private void сохранитьТаблицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (B == false)
            {
                MessageBox.Show("Сначала нужно создать/загрузить таблицу","Ошибка");
                return;
            }
            else 
            {
            SAVE gekus = new SAVE();
            gekus.Main1(dataGridView1);
            }
        }
        public class SAVE //СОХРАНЕНИЕ ТАБЛИЦЫ
        {
            
            public void Main1(DataGridView dataGridView1)
            {
                int row_count = dataGridView1.RowCount;
                int col_count = dataGridView1.ColumnCount;

                string[,] x = new string[row_count, col_count];

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++) {
                        if (dataGridView1.Rows[i].Cells[j].Value == null) {
                            x[i, j] = " ";
                        }
                        else x[i, j] = dataGridView1.Rows[i].Cells[j].Value.ToString(); 
                    }
                }
                string newExcelFile = @"C:\Users\Maksim\Desktop\Расписание.xlsx";
                new SAVE().Export(newExcelFile, x);

            }
            public void Export(string file, string[,] x)
            {
                x[44, 16] = "Конец файла";
                var list = new List<Person>

            {
            new Person {День_недели= x[0,0], Урок = x[0,1], _5А_ = x[0,2],_5Б_ = x[0,3], _5В_ = x[0,4], _6А_=x[0,5], _6Б_= x[0,6], _6В_=x[0,7], _7А_=x[0,8],_7Б_=x[0,9],_7В_=x[0,10],_8А_=x[0,11],_8Б_=x[0,12],_8В_=x[0,13],_9А_ = x[0,14],_9Б_ = x[0,15], _9В_=x[0,16], Запасное=x[0,17]},
            new Person {День_недели= x[1,0], Урок = x[1,1], _5А_ = x[1,2],_5Б_ = x[1,3], _5В_ = x[1,4], _6А_=x[1,5], _6Б_= x[1,6], _6В_=x[1,7], _7А_=x[1,8],_7Б_=x[1,9],_7В_=x[1,10],_8А_=x[1,11],_8Б_=x[1,12],_8В_=x[1,13],_9А_ = x[1,14],_9Б_ = x[1,15], _9В_=x[1,16], Запасное=x[1,17]},
            new Person {День_недели= x[2,0], Урок = x[2,1], _5А_ = x[2,2],_5Б_ = x[2,3], _5В_ = x[2,4], _6А_=x[2,5], _6Б_= x[2,6], _6В_=x[2,7], _7А_=x[2,8],_7Б_=x[2,9],_7В_=x[2,10],_8А_=x[2,11],_8Б_=x[2,12],_8В_=x[2,13],_9А_ = x[2,14],_9Б_ = x[2,15], _9В_=x[2,16], Запасное=x[2,17]},
            new Person {День_недели= x[3,0], Урок = x[3,1], _5А_ = x[3,2],_5Б_ = x[3,3], _5В_ = x[3,4], _6А_=x[3,5], _6Б_= x[3,6], _6В_=x[3,7], _7А_=x[3,8],_7Б_=x[3,9],_7В_=x[3,10],_8А_=x[3,11],_8Б_=x[3,12],_8В_=x[3,13],_9А_ = x[3,14],_9Б_ = x[3,15], _9В_=x[3,16], Запасное=x[3,17]},
            new Person {День_недели= x[4,0], Урок = x[4,1], _5А_ = x[4,2],_5Б_ = x[4,3], _5В_ = x[4,4], _6А_=x[4,5], _6Б_= x[4,6], _6В_=x[4,7], _7А_=x[4,8],_7Б_=x[4,9],_7В_=x[4,10],_8А_=x[4,11],_8Б_=x[4,12],_8В_=x[4,13],_9А_ = x[4,14],_9Б_ = x[4,15], _9В_=x[4,16], Запасное=x[4,17]},
            new Person {День_недели= x[5,0], Урок = x[5,1], _5А_ = x[5,2],_5Б_ = x[5,3], _5В_ = x[5,4], _6А_=x[5,5], _6Б_= x[5,6], _6В_=x[5,7], _7А_=x[5,8],_7Б_=x[5,9],_7В_=x[5,10],_8А_=x[5,11],_8Б_=x[5,12],_8В_=x[5,13],_9А_ = x[5,14],_9Б_ = x[5,15], _9В_=x[5,16], Запасное=x[5,17]},
            new Person {День_недели= x[6,0], Урок = x[6,1], _5А_ = x[6,2],_5Б_ = x[6,3], _5В_ = x[6,4], _6А_=x[6,5], _6Б_= x[6,6], _6В_=x[6,7], _7А_=x[6,8],_7Б_=x[6,9],_7В_=x[6,10],_8А_=x[6,11],_8Б_=x[6,12],_8В_=x[6,13],_9А_ = x[6,14],_9Б_ = x[6,15], _9В_=x[6,16], Запасное=x[6,17]},
            new Person {День_недели= x[7,0], Урок = x[7,1], _5А_ = x[7,2],_5Б_ = x[7,3], _5В_ = x[7,4], _6А_=x[7,5], _6Б_= x[7,6], _6В_=x[7,7], _7А_=x[7,8],_7Б_=x[7,9],_7В_=x[7,10],_8А_=x[7,11],_8Б_=x[7,12],_8В_=x[7,13],_9А_ = x[7,14],_9Б_ = x[7,15], _9В_=x[7,16], Запасное=x[7,17]},
            new Person {День_недели= x[8,0], Урок = x[8,1], _5А_ = x[8,2],_5Б_ = x[8,3], _5В_ = x[8,4], _6А_=x[8,5], _6Б_= x[8,6], _6В_=x[8,7], _7А_=x[8,8],_7Б_=x[8,9],_7В_=x[8,10],_8А_=x[8,11],_8Б_=x[8,12],_8В_=x[8,13],_9А_ = x[8,14],_9Б_ = x[8,15], _9В_=x[8,16], Запасное=x[8,17]},
            new Person {День_недели= x[9,0], Урок = x[9,1], _5А_ = x[9,2],_5Б_ = x[9,3], _5В_ = x[9,4], _6А_=x[9,5], _6Б_= x[9,6], _6В_=x[9,7], _7А_=x[9,8],_7Б_=x[9,9],_7В_=x[9,10],_8А_=x[9,11],_8Б_=x[9,12],_8В_=x[9,13],_9А_ = x[9,14],_9Б_ = x[9,15], _9В_=x[9,16], Запасное=x[9,17]},
            new Person {День_недели= x[10,0], Урок = x[10,1], _5А_ = x[10,2],_5Б_ = x[10,3], _5В_ = x[10,4], _6А_=x[10,5], _6Б_= x[10,6], _6В_=x[10,7], _7А_=x[10,8],_7Б_=x[10,9],_7В_=x[10,10],_8А_=x[10,11],_8Б_=x[10,12],_8В_=x[10,13],_9А_ = x[10,14],_9Б_ = x[10,15], _9В_=x[10,16], Запасное=x[10,17]},
            new Person {День_недели= x[11,0], Урок = x[11,1], _5А_ = x[11,2],_5Б_ = x[11,3], _5В_ = x[11,4], _6А_=x[11,5], _6Б_= x[11,6], _6В_=x[11,7], _7А_=x[11,8],_7Б_=x[11,9],_7В_=x[11,10],_8А_=x[11,11],_8Б_=x[11,12],_8В_=x[11,13],_9А_ = x[11,14],_9Б_ = x[11,15], _9В_=x[11,16], Запасное=x[11,17]},
            new Person {День_недели= x[12,0], Урок = x[12,1], _5А_ = x[12,2],_5Б_ = x[12,3], _5В_ = x[12,4], _6А_=x[12,5], _6Б_= x[12,6], _6В_=x[12,7], _7А_=x[12,8],_7Б_=x[12,9],_7В_=x[12,10],_8А_=x[12,11],_8Б_=x[12,12],_8В_=x[12,13],_9А_ = x[12,14],_9Б_ = x[12,15], _9В_=x[12,16], Запасное=x[12,17]},
            new Person {День_недели= x[13,0], Урок = x[13,1], _5А_ = x[13,2],_5Б_ = x[13,3], _5В_ = x[13,4], _6А_=x[13,5], _6Б_= x[13,6], _6В_=x[13,7], _7А_=x[13,8],_7Б_=x[13,9],_7В_=x[13,10],_8А_=x[13,11],_8Б_=x[13,12],_8В_=x[13,13],_9А_ = x[13,14],_9Б_ = x[13,15], _9В_=x[13,16], Запасное=x[13,17]},
            new Person {День_недели= x[14,0], Урок = x[14,1], _5А_ = x[14,2],_5Б_ = x[14,3], _5В_ = x[14,4], _6А_=x[14,5], _6Б_= x[14,6], _6В_=x[14,7], _7А_=x[14,8],_7Б_=x[14,9],_7В_=x[14,10],_8А_=x[14,11],_8Б_=x[14,12],_8В_=x[14,13],_9А_ = x[14,14],_9Б_ = x[14,15], _9В_=x[14,16], Запасное=x[14,17]},
            new Person {День_недели= x[15,0], Урок = x[15,1], _5А_ = x[15,2],_5Б_ = x[15,3], _5В_ = x[15,4], _6А_=x[15,5], _6Б_= x[15,6], _6В_=x[15,7], _7А_=x[15,8],_7Б_=x[15,9],_7В_=x[15,10],_8А_=x[15,11],_8Б_=x[15,12],_8В_=x[15,13],_9А_ = x[15,14],_9Б_ = x[15,15], _9В_=x[15,16], Запасное=x[15,17]},
            new Person {День_недели= x[16,0], Урок = x[16,1], _5А_ = x[16,2],_5Б_ = x[16,3], _5В_ = x[16,4], _6А_=x[16,5], _6Б_= x[16,6], _6В_=x[16,7], _7А_=x[16,8],_7Б_=x[16,9],_7В_=x[16,10],_8А_=x[16,11],_8Б_=x[16,12],_8В_=x[16,13],_9А_ = x[16,14],_9Б_ = x[16,15], _9В_=x[16,16], Запасное=x[16,17]},
            new Person {День_недели= x[17,0], Урок = x[17,1], _5А_ = x[17,2],_5Б_ = x[17,3], _5В_ = x[17,4], _6А_=x[17,5], _6Б_= x[17,6], _6В_=x[17,7], _7А_=x[17,8],_7Б_=x[17,9],_7В_=x[17,10],_8А_=x[17,11],_8Б_=x[17,12],_8В_=x[17,13],_9А_ = x[17,14],_9Б_ = x[17,15], _9В_=x[17,16], Запасное=x[17,17]},
            new Person {День_недели= x[18,0], Урок = x[18,1], _5А_ = x[18,2],_5Б_ = x[18,3], _5В_ = x[18,4], _6А_=x[18,5], _6Б_= x[18,6], _6В_=x[18,7], _7А_=x[18,8],_7Б_=x[18,9],_7В_=x[18,10],_8А_=x[18,11],_8Б_=x[18,12],_8В_=x[18,13],_9А_ = x[18,14],_9Б_ = x[18,15], _9В_=x[18,16], Запасное=x[18,17]},
            new Person {День_недели= x[19,0], Урок = x[19,1], _5А_ = x[19,2],_5Б_ = x[19,3], _5В_ = x[19,4], _6А_=x[19,5], _6Б_= x[19,6], _6В_=x[19,7], _7А_=x[19,8],_7Б_=x[19,9],_7В_=x[19,10],_8А_=x[19,11],_8Б_=x[19,12],_8В_=x[19,13],_9А_ = x[19,14],_9Б_ = x[19,15], _9В_=x[19,16], Запасное=x[19,17]},
            new Person {День_недели= x[20,0], Урок = x[20,1], _5А_ = x[20,2],_5Б_ = x[20,3], _5В_ = x[20,4], _6А_=x[20,5], _6Б_= x[20,6], _6В_=x[20,7], _7А_=x[20,8],_7Б_=x[20,9],_7В_=x[20,10],_8А_=x[20,11],_8Б_=x[20,12],_8В_=x[20,13],_9А_ = x[20,14],_9Б_ = x[20,15], _9В_=x[20,16], Запасное=x[20,17]},
            new Person {День_недели= x[21,0], Урок = x[21,1], _5А_ = x[21,2],_5Б_ = x[21,3], _5В_ = x[21,4], _6А_=x[21,5], _6Б_= x[21,6], _6В_=x[21,7], _7А_=x[21,8],_7Б_=x[21,9],_7В_=x[21,10],_8А_=x[21,11],_8Б_=x[21,12],_8В_=x[21,13],_9А_ = x[21,14],_9Б_ = x[21,15], _9В_=x[21,16], Запасное=x[21,17]},
            new Person {День_недели= x[22,0], Урок = x[22,1], _5А_ = x[22,2],_5Б_ = x[22,3], _5В_ = x[22,4], _6А_=x[22,5], _6Б_= x[22,6], _6В_=x[22,7], _7А_=x[22,8],_7Б_=x[22,9],_7В_=x[22,10],_8А_=x[22,11],_8Б_=x[22,12],_8В_=x[22,13],_9А_ = x[22,14],_9Б_ = x[22,15], _9В_=x[22,16], Запасное=x[22,17]},
            new Person {День_недели= x[23,0], Урок = x[23,1], _5А_ = x[23,2],_5Б_ = x[23,3], _5В_ = x[23,4], _6А_=x[23,5], _6Б_= x[23,6], _6В_=x[23,7], _7А_=x[23,8],_7Б_=x[23,9],_7В_=x[23,10],_8А_=x[23,11],_8Б_=x[23,12],_8В_=x[23,13],_9А_ = x[23,14],_9Б_ = x[23,15], _9В_=x[23,16], Запасное=x[23,17]},
            new Person {День_недели= x[24,0], Урок = x[24,1], _5А_ = x[24,2],_5Б_ = x[24,3], _5В_ = x[24,4], _6А_=x[24,5], _6Б_= x[24,6], _6В_=x[24,7], _7А_=x[24,8],_7Б_=x[24,9],_7В_=x[24,10],_8А_=x[24,11],_8Б_=x[24,12],_8В_=x[24,13],_9А_ = x[24,14],_9Б_ = x[24,15], _9В_=x[24,16], Запасное=x[24,17]},
            new Person {День_недели= x[25,0], Урок = x[25,1], _5А_ = x[25,2],_5Б_ = x[25,3], _5В_ = x[25,4], _6А_=x[25,5], _6Б_= x[25,6], _6В_=x[25,7], _7А_=x[25,8],_7Б_=x[25,9],_7В_=x[25,10],_8А_=x[25,11],_8Б_=x[25,12],_8В_=x[25,13],_9А_ = x[25,14],_9Б_ = x[25,15], _9В_=x[25,16], Запасное=x[25,17]},
            new Person {День_недели= x[26,0], Урок = x[26,1], _5А_ = x[26,2],_5Б_ = x[26,3], _5В_ = x[26,4], _6А_=x[26,5], _6Б_= x[26,6], _6В_=x[26,7], _7А_=x[26,8],_7Б_=x[26,9],_7В_=x[26,10],_8А_=x[26,11],_8Б_=x[26,12],_8В_=x[26,13],_9А_ = x[26,14],_9Б_ = x[26,15], _9В_=x[26,16], Запасное=x[26,17]},
            new Person {День_недели= x[27,0], Урок = x[27,1], _5А_ = x[27,2],_5Б_ = x[27,3], _5В_ = x[27,4], _6А_=x[27,5], _6Б_= x[27,6], _6В_=x[27,7], _7А_=x[27,8],_7Б_=x[27,9],_7В_=x[27,10],_8А_=x[27,11],_8Б_=x[27,12],_8В_=x[27,13],_9А_ = x[27,14],_9Б_ = x[27,15], _9В_=x[27,16], Запасное=x[27,17]},
            new Person {День_недели= x[28,0], Урок = x[28,1], _5А_ = x[28,2],_5Б_ = x[28,3], _5В_ = x[28,4], _6А_=x[28,5], _6Б_= x[28,6], _6В_=x[28,7], _7А_=x[28,8],_7Б_=x[28,9],_7В_=x[28,10],_8А_=x[28,11],_8Б_=x[28,12],_8В_=x[28,13],_9А_ = x[28,14],_9Б_ = x[28,15], _9В_=x[28,16], Запасное=x[28,17]},
            new Person {День_недели= x[29,0], Урок = x[29,1], _5А_ = x[29,2],_5Б_ = x[29,3], _5В_ = x[29,4], _6А_=x[29,5], _6Б_= x[29,6], _6В_=x[29,7], _7А_=x[29,8],_7Б_=x[29,9],_7В_=x[29,10],_8А_=x[29,11],_8Б_=x[29,12],_8В_=x[29,13],_9А_ = x[29,14],_9Б_ = x[29,15], _9В_=x[29,16], Запасное=x[29,17]},
            new Person {День_недели= x[30,0], Урок = x[30,1], _5А_ = x[30,2],_5Б_ = x[30,3], _5В_ = x[30,4], _6А_=x[30,5], _6Б_= x[30,6], _6В_=x[30,7], _7А_=x[30,8],_7Б_=x[30,9],_7В_=x[30,10],_8А_=x[30,11],_8Б_=x[30,12],_8В_=x[30,13],_9А_ = x[30,14],_9Б_ = x[30,15], _9В_=x[30,16], Запасное=x[30,17]},
            new Person {День_недели= x[31,0], Урок = x[31,1], _5А_ = x[31,2],_5Б_ = x[31,3], _5В_ = x[31,4], _6А_=x[31,5], _6Б_= x[31,6], _6В_=x[31,7], _7А_=x[31,8],_7Б_=x[31,9],_7В_=x[31,10],_8А_=x[31,11],_8Б_=x[31,12],_8В_=x[31,13],_9А_ = x[31,14],_9Б_ = x[31,15], _9В_=x[31,16], Запасное=x[31,17]},
            new Person {День_недели= x[32,0], Урок = x[32,1], _5А_ = x[32,2],_5Б_ = x[32,3], _5В_ = x[32,4], _6А_=x[32,5], _6Б_= x[32,6], _6В_=x[32,7], _7А_=x[32,8],_7Б_=x[32,9],_7В_=x[32,10],_8А_=x[32,11],_8Б_=x[32,12],_8В_=x[32,13],_9А_ = x[32,14],_9Б_ = x[32,15], _9В_=x[32,16], Запасное=x[32,17]},
            new Person {День_недели= x[33,0], Урок = x[33,1], _5А_ = x[33,2],_5Б_ = x[33,3], _5В_ = x[33,4], _6А_=x[33,5], _6Б_= x[33,6], _6В_=x[33,7], _7А_=x[33,8],_7Б_=x[33,9],_7В_=x[33,10],_8А_=x[33,11],_8Б_=x[33,12],_8В_=x[33,13],_9А_ = x[33,14],_9Б_ = x[33,15], _9В_=x[33,16], Запасное=x[33,17]},
            new Person {День_недели= x[34,0], Урок = x[34,1], _5А_ = x[34,2],_5Б_ = x[34,3], _5В_ = x[34,4], _6А_=x[34,5], _6Б_= x[34,6], _6В_=x[34,7], _7А_=x[34,8],_7Б_=x[34,9],_7В_=x[34,10],_8А_=x[34,11],_8Б_=x[34,12],_8В_=x[34,13],_9А_ = x[34,14],_9Б_ = x[34,15], _9В_=x[34,16], Запасное=x[34,17]},
            new Person {День_недели= x[35,0], Урок = x[35,1], _5А_ = x[35,2],_5Б_ = x[35,3], _5В_ = x[35,4], _6А_=x[35,5], _6Б_= x[35,6], _6В_=x[35,7], _7А_=x[35,8],_7Б_=x[35,9],_7В_=x[35,10],_8А_=x[35,11],_8Б_=x[35,12],_8В_=x[35,13],_9А_ = x[35,14],_9Б_ = x[35,15], _9В_=x[35,16], Запасное=x[35,17]},
            new Person {День_недели= x[36,0], Урок = x[36,1], _5А_ = x[36,2],_5Б_ = x[36,3], _5В_ = x[36,4], _6А_=x[36,5], _6Б_= x[36,6], _6В_=x[36,7], _7А_=x[36,8],_7Б_=x[36,9],_7В_=x[36,10],_8А_=x[36,11],_8Б_=x[36,12],_8В_=x[36,13],_9А_ = x[36,14],_9Б_ = x[36,15], _9В_=x[36,16], Запасное=x[36,17]},
            new Person {День_недели= x[37,0], Урок = x[37,1], _5А_ = x[37,2],_5Б_ = x[37,3], _5В_ = x[37,4], _6А_=x[37,5], _6Б_= x[37,6], _6В_=x[37,7], _7А_=x[37,8],_7Б_=x[37,9],_7В_=x[37,10],_8А_=x[37,11],_8Б_=x[37,12],_8В_=x[37,13],_9А_ = x[37,14],_9Б_ = x[37,15], _9В_=x[37,16], Запасное=x[37,17]},
            new Person {День_недели= x[38,0], Урок = x[38,1], _5А_ = x[38,2],_5Б_ = x[38,3], _5В_ = x[38,4], _6А_=x[38,5], _6Б_= x[38,6], _6В_=x[38,7], _7А_=x[38,8],_7Б_=x[38,9],_7В_=x[38,10],_8А_=x[38,11],_8Б_=x[38,12],_8В_=x[38,13],_9А_ = x[38,14],_9Б_ = x[38,15], _9В_=x[38,16], Запасное=x[38,17]},
            new Person {День_недели= x[39,0], Урок = x[39,1], _5А_ = x[39,2],_5Б_ = x[39,3], _5В_ = x[39,4], _6А_=x[39,5], _6Б_= x[39,6], _6В_=x[39,7], _7А_=x[39,8],_7Б_=x[39,9],_7В_=x[39,10],_8А_=x[39,11],_8Б_=x[39,12],_8В_=x[39,13],_9А_ = x[39,14],_9Б_ = x[39,15], _9В_=x[39,16], Запасное=x[39,17]},
            new Person {День_недели= x[40,0], Урок = x[40,1], _5А_ = x[40,2],_5Б_ = x[40,3], _5В_ = x[40,4], _6А_=x[40,5], _6Б_= x[40,6], _6В_=x[40,7], _7А_=x[40,8],_7Б_=x[40,9],_7В_=x[40,10],_8А_=x[40,11],_8Б_=x[40,12],_8В_=x[40,13],_9А_ = x[40,14],_9Б_ = x[40,15], _9В_=x[40,16], Запасное=x[40,17]},
            new Person {День_недели= x[41,0], Урок = x[41,1], _5А_ = x[41,2],_5Б_ = x[41,3], _5В_ = x[41,4], _6А_=x[41,5], _6Б_= x[41,6], _6В_=x[41,7], _7А_=x[41,8],_7Б_=x[41,9],_7В_=x[41,10],_8А_=x[41,11],_8Б_=x[41,12],_8В_=x[41,13],_9А_ = x[41,14],_9Б_ = x[41,15], _9В_=x[41,16], Запасное=x[41,17]},
            new Person {День_недели= x[42,0], Урок = x[42,1], _5А_ = x[42,2],_5Б_ = x[42,3], _5В_ = x[42,4], _6А_=x[42,5], _6Б_= x[42,6], _6В_=x[42,7], _7А_=x[42,8],_7Б_=x[42,9],_7В_=x[42,10],_8А_=x[42,11],_8Б_=x[42,12],_8В_=x[42,13],_9А_ = x[42,14],_9Б_ = x[42,15], _9В_=x[42,16], Запасное=x[42,17]},
            new Person {День_недели= x[43,0], Урок = x[43,1], _5А_ = x[43,2],_5Б_ = x[43,3], _5В_ = x[43,4], _6А_=x[43,5], _6Б_= x[43,6], _6В_=x[43,7], _7А_=x[43,8],_7Б_=x[43,9],_7В_=x[43,10],_8А_=x[43,11],_8Б_=x[43,12],_8В_=x[43,13],_9А_ = x[43,14],_9Б_ = x[43,15], _9В_=x[43,16], Запасное=x[43,17]},
            new Person {День_недели= x[44,0], Урок = x[44,1], _5А_ = x[44,2],_5Б_ = x[44,3], _5В_ = x[44,4], _6А_=x[44,5], _6Б_= x[44,6], _6В_=x[44,7], _7А_=x[43,8],_7Б_=x[44,9],_7В_=x[44,10],_8А_=x[44,11],_8Б_=x[44,12],_8В_=x[44,13],_9А_ = x[44,14],_9Б_ = x[44,15], _9В_=x[44,16], Запасное=x[44,17]},
            };
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage pck = new ExcelPackage())
                {
                    pck.Workbook.Worksheets.Add("Расписание").Cells[1, 1].LoadFromCollection(list, true);
                    pck.SaveAs(new FileInfo(file));
                }
            }
            public class Person
            {
                public string День_недели { get; set; }
                public string Урок { get; set; }
                public string _5А_ { get; set; }
                public string _5Б_ { get; set; }
                public string _5В_ { get; set; }
                public string _6А_ { get; set; }
                public string _6Б_ { get; set; }
                public string _6В_ { get; set; }
                public string _7А_ { get; set; }
                public string _7Б_ { get; set; }
                public string _7В_ { get; set; }
                public string _8А_ { get; set; }
                public string _8Б_ { get; set; }
                public string _8В_ { get; set; }
                public string _9А_ { get; set; }
                public string _9Б_ { get; set; }
                public string _9В_ { get; set; }
                public string Запасное { get; set; }
            }
        }

        private void загрузитьТаблицуToolStripMenuItem_Click(object sender, EventArgs e) //ОТКРЫТЬ СУЩЕСТВУЮЩЮЮ ТАБЛИЦУ
        {
            V = true;
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        this.dataGridView1.Columns.RemoveAt(i);
                    };
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        this.dataGridView1.Columns.RemoveAt(i);
                    };
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        this.dataGridView1.Columns.RemoveAt(i);
                    };
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        this.dataGridView1.Columns.RemoveAt(i);
                    };
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        this.dataGridView1.Columns.RemoveAt(i);
                    };
                    dataGridView1.Refresh();
                    fileName = openFileDialog1.FileName;
                    Text = fileName;
                    OpenExcelFile(fileName);
                }
                else
                {
                    B = false;
                    throw new Exception("Файл не выбран");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            tableCollection = db.Tables;
            toolStripComboBox1.Items.Clear();
            foreach (DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
            toolStripComboBox1.Visible = true;
            toolStripLabel1.Visible = true;
            dataGridView1.Rows[0].Cells[0].Selected = false;
            B = true;
            
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {
            if (V == true)
            {
                DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
                dataGridView1.DataSource = table;
            }
        }

        private void сохранитьToolStripButton_Click(object sender, EventArgs e)
        {
            if (B == false)
            {
                MessageBox.Show("Нечего сохранять", "Ошибка");
                return;
            }
            else
            {
                SAVE gekus = new SAVE();
                gekus.Main1(dataGridView1);
            }
        }

        private void сгенерироватьМакетToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (B==false)
                MessageBox.Show("Таблица не создана", "Ошибка");
            else
            {
                dataGridView1.Rows[0].Cells[0].Value = "День недели";                dataGridView1.Rows[0].Cells[1].Value = "Время урока";                dataGridView1.Rows[1].Cells[0].Value = "Понедельник";
                dataGridView1.Rows[1].Cells[1].Value = "8:00-8:40";                dataGridView1.Rows[2].Cells[1].Value = "8:50-9:30";                dataGridView1.Rows[3].Cells[1].Value = "9:50-10:30";
                dataGridView1.Rows[4].Cells[1].Value = "10:50-11:30";                dataGridView1.Rows[5].Cells[1].Value = "11:50-12:30";                dataGridView1.Rows[6].Cells[1].Value = "12:40-13:20";
                dataGridView1.Rows[7].Cells[1].Value = "13:30-14:10";                dataGridView1.Rows[8].Cells[0].Value = "Вторник";                dataGridView1.Rows[8].Cells[1].Value = "8:00-8:40";
                dataGridView1.Rows[9].Cells[1].Value = "8:50-9:30";                dataGridView1.Rows[10].Cells[1].Value = "9:50-10:30";                dataGridView1.Rows[11].Cells[1].Value = "10:50-11:30";
                dataGridView1.Rows[12].Cells[1].Value = "11:50-12:30";                dataGridView1.Rows[13].Cells[1].Value = "12:40-13:20";                dataGridView1.Rows[14].Cells[1].Value = "13:30-14:10";
                dataGridView1.Rows[15].Cells[0].Value = "Среда";                dataGridView1.Rows[15].Cells[1].Value = "8:00-8:40";                dataGridView1.Rows[16].Cells[1].Value = "8:50-9:30";
                dataGridView1.Rows[17].Cells[1].Value = "9:50-10:30";                dataGridView1.Rows[18].Cells[1].Value = "10:50-11:30";                dataGridView1.Rows[19].Cells[1].Value = "11:50-12:30";
                dataGridView1.Rows[20].Cells[1].Value = "12:40-13:20";                dataGridView1.Rows[21].Cells[1].Value = "13:30-14:10";                dataGridView1.Rows[22].Cells[0].Value = "Четверг";
                dataGridView1.Rows[22].Cells[1].Value = "8:00-8:40";                dataGridView1.Rows[23].Cells[1].Value = "8:50-9:30";                dataGridView1.Rows[24].Cells[1].Value = "9:50-10:30";
                dataGridView1.Rows[25].Cells[1].Value = "10:50-11:30";                dataGridView1.Rows[26].Cells[1].Value = "11:50-12:30";                dataGridView1.Rows[27].Cells[1].Value = "12:40-13:20";
                dataGridView1.Rows[28].Cells[1].Value = "13:30-14:10";                dataGridView1.Rows[29].Cells[0].Value = "Пятница";                dataGridView1.Rows[29].Cells[1].Value = "8:00-8:40";
                dataGridView1.Rows[30].Cells[1].Value = "8:50-9:30";                dataGridView1.Rows[31].Cells[1].Value = "9:50-10:30";                dataGridView1.Rows[32].Cells[1].Value = "10:50-11:30";
                dataGridView1.Rows[33].Cells[1].Value = "11:50-12:30";                dataGridView1.Rows[34].Cells[1].Value = "12:40-13:20";                dataGridView1.Rows[35].Cells[1].Value = "13:30-14:10";
                dataGridView1.Rows[0].Cells[2].Value = "5А";                dataGridView1.Rows[0].Cells[3].Value = "5Б";                dataGridView1.Rows[0].Cells[4].Value = "5В";
                dataGridView1.Rows[0].Cells[5].Value = "6А";                dataGridView1.Rows[0].Cells[6].Value = "6Б";                dataGridView1.Rows[0].Cells[7].Value = "6В";
                dataGridView1.Rows[0].Cells[8].Value = "7А";                dataGridView1.Rows[0].Cells[9].Value = "7Б";                dataGridView1.Rows[0].Cells[10].Value = "7В";
                dataGridView1.Rows[0].Cells[11].Value = "8А";                dataGridView1.Rows[0].Cells[12].Value = "8Б";                dataGridView1.Rows[0].Cells[13].Value = "8В";
                dataGridView1.Rows[0].Cells[14].Value = "9А";                dataGridView1.Rows[0].Cells[15].Value = "9Б";                dataGridView1.Rows[0].Cells[16].Value = "9В";

            }
                
        }

        private void справкаToolStripButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Программа предназначена для создания и редактирования расписания.\nДля работы с программой необходимо создать или загрузить таблицу.\n После создания таблицы можно выполнить редактирование\nПрограмму создал студент группы 081 Гекк М.Е. ","О программе");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                            if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                            {
                                dataGridView1.Rows[i].Cells[j].Selected = true;
                            }
                }
            }
        }
    }
} 

