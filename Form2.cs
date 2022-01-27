using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Data.SqlClient;

namespace project1_Parser
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd. MM. yyyy г., HH:m";
            dateTimePicker1.Value = DateTime.Now;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            dateTimePicker1.Enabled = false;
            textBox2.Enabled = false;
            button3.Enabled = false;
            button6.Enabled = false;

        }
        static List<string> docxTableGroup = new List<string> { };
        static List<string> docxTableStudent = new List<string> { };
        static List<string> docxTableLesson = new List<string> { };
        static List<string> auditoriyaPeresdachi = new List<string> { };
        static List<DateTime> vremyaPeresdachi = new List<DateTime> { };
        static string documentPut;
        static string teachers;

        string IdDepartmentSelect(string full_name_department)
        {
            string id_department;
            SqlCommand command1 = new SqlCommand($"SELECT [id_department] FROM [departments] WHERE [full_name_department] = N'{full_name_department}'", Program.sqlConnection);
            command1.Parameters.AddWithValue("full_name_department", full_name_department);
            
            id_department = Convert.ToString(command1.ExecuteScalar());
            
            return id_department;
        }

        string NameDepartmentSelect(string full_name_department)
        {
            string fn_department;
            SqlCommand command1 = new SqlCommand($"SELECT [name_department] FROM [departments] WHERE [full_name_department] = N'{full_name_department}'", Program.sqlConnection);
            command1.Parameters.AddWithValue("full_name_department", full_name_department);
            
            fn_department = Convert.ToString(command1.ExecuteScalar());
            
            return fn_department;
        }

        void FillingComboTeachers(string id_department)
        {
            SqlCommand commandSelect = new SqlCommand($"SELECT [name_teachers] FROM [teachers] WHERE [id_department] = N'{id_department}' ORDER BY [name_teachers]", Program.sqlConnection);
            commandSelect.Parameters.AddWithValue("id_department", id_department);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = commandSelect.ExecuteReader();

                while (sqlReader.Read())
                {
                    checkedListBox2.Items.Add(Convert.ToString(sqlReader["name_teachers"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
                Program.sqlConnection.Close();
            }
        }


        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private async void Button1_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                documentPut = saveFileDialog1.FileName;
                comboBox1.Enabled = true;
                comboBox1.Visible = true;
            }
            comboBox1.Items.Clear();
            Program.Connect();
            await Program.sqlConnection.OpenAsync();
            ComboInstitutes();

        }

        private void Shapka_docx (object sender, EventArgs e)
        {
            DocX document = DocX.Create(documentPut);
            document.MarginBottom = 25;
            document.MarginLeft = 50;
            document.MarginRight = 50;
            document.MarginTop = 25;
            document.InsertParagraph("УТВЕРЖДАЮ            ").Font("Times New Roman").FontSize(12).Bold().IndentationFirstLine = 360;
            document.InsertParagraph("директор ИЦС         ").Font("Times New Roman").FontSize(12).IndentationFirstLine = 360;
            document.InsertParagraph("_____Т.К. Ивашковская").Font("Times New Roman").FontSize(12).IndentationFirstLine = 360;
            document.InsertParagraph("«__»__________2021 г.").Font("Times New Roman").FontSize(12).IndentationFirstLine = 360;

            document.InsertParagraph("").Font("Times New Roman").FontSize(14).Bold().Alignment = Xceed.Document.NET.Alignment.center;

            document.InsertParagraph("ГРАФИК").Font("Times New Roman").FontSize(14).Bold().Alignment = Xceed.Document.NET.Alignment.center;
            document.InsertParagraph("проведения второй повторной промежуточной аттестации для студентов").Font("Times New Roman").FontSize(14).Alignment = Xceed.Document.NET.Alignment.center;
            document.InsertParagraph("института цифровых систем, имеющих задолженности по дисциплинам кафедры").Font("Times New Roman").FontSize(14).Alignment = Xceed.Document.NET.Alignment.center;
            document.InsertParagraph("«" + comboBox1.Text + "»").Font("Times New Roman").FontSize(14).Alignment = Xceed.Document.NET.Alignment.center;

            document.Save();
        }


        private void Button3_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            comboBox1.Enabled = false;
            Shapka_docx(sender, e);
            comboBox2.Enabled = true;
            dateTimePicker1.Enabled = true;
            textBox2.Enabled = true;
        }

        private void ComboBox1_TextUpdate(object sender, EventArgs e)
        {
            
        }

        private void ComboBox2_TextUpdate(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                dateTimePicker1.Enabled = true;
                textBox2.Enabled = true;
            }
            else
            {
                dateTimePicker1.Enabled = false;
                textBox2.Enabled = false;
                button6.Enabled = false;
            }

        }

        private void DocxTableCreator(object sender, EventArgs e)
        {
            DocX document = null;
            if (documentPut.Contains(".docx"))
            {
                document = DocX.Load(documentPut);
            }
            else
            {
                document = DocX.Load(documentPut + ".docx");
            }
            
            Table table = document.InsertTable(docxTableStudent.Count + 1, 3);
            table.Alignment = Alignment.left;
            table.Rows[0].Cells[0].RemoveParagraphAt(0);
            table.Rows[0].Cells[1].RemoveParagraphAt(0);
            table.Rows[0].Cells[2].RemoveParagraphAt(0);
            table.Rows[0].Cells[0].InsertParagraph("Группа").FontSize(12).Alignment = Alignment.center;
            table.Rows[0].Cells[1].InsertParagraph("ФИО студента").FontSize(12).Alignment = Alignment.center;
            table.Rows[0].Cells[2].InsertParagraph("Дисциплина").FontSize(12).Alignment = Alignment.center;

            if (docxTableStudent.Count > 1)
            {
                int ttemp = 0;
                string temp1, temp2, temp3;
                for (int j = 0; j < docxTableStudent.Count - 1; j++)
                {
                    for (int i = j + 1; i < docxTableStudent.Count; i++)
                    {
                        if (docxTableLesson[ttemp] == docxTableLesson[i])
                        {
                            temp1 = docxTableGroup[ttemp + 1];
                            temp2 = docxTableStudent[ttemp + 1];
                            temp3 = docxTableLesson[ttemp + 1];
                            docxTableGroup[ttemp + 1] = docxTableGroup[i];
                            docxTableStudent[ttemp + 1] = docxTableStudent[i];
                            docxTableLesson[ttemp + 1] = docxTableLesson[i];
                            docxTableGroup[i] = temp1;
                            docxTableStudent[i] = temp2;
                            docxTableLesson[i] = temp3;
                            j++;
                            ttemp++;
                        }
                    }
                    ttemp++;
                }

                ttemp = 0;

                for (int j = 0; j < docxTableStudent.Count - 1; j++)
                {
                    for (int i = j + 1; i < docxTableStudent.Count; i++)
                    {
                        if (docxTableGroup[ttemp] == docxTableGroup[i] && docxTableLesson[ttemp] == docxTableLesson[i])
                        {
                            temp1 = docxTableGroup[ttemp + 1];
                            temp2 = docxTableStudent[ttemp + 1];
                            temp3 = docxTableLesson[ttemp + 1];
                            docxTableGroup[ttemp + 1] = docxTableGroup[i];
                            docxTableStudent[ttemp + 1] = docxTableStudent[i];
                            docxTableLesson[ttemp + 1] = docxTableLesson[i];
                            docxTableGroup[i] = temp1;
                            docxTableStudent[i] = temp2;
                            docxTableLesson[i] = temp3;
                            j++;
                            ttemp++;
                        }
                    }
                    ttemp++;
                }

                ttemp = 0;

                for (int row = 1; row < docxTableStudent.Count + 1; row++)
                {
                    table.Rows[row].Cells[0].RemoveParagraphAt(0);
                    table.Rows[row].Cells[1].RemoveParagraphAt(0);
                    table.Rows[row].Cells[2].RemoveParagraphAt(0);
                    Paragraph cell_paragraph = table.Rows[row].Cells[0].InsertParagraph(docxTableGroup[row - 1]).FontSize(12);
                    cell_paragraph.Alignment = Alignment.center;

                    cell_paragraph = table.Rows[row].Cells[1].InsertParagraph(docxTableStudent[row - 1]).FontSize(12);
                    cell_paragraph.Alignment = Alignment.center;

                    cell_paragraph = table.Rows[row].Cells[2].InsertParagraph(docxTableLesson[row - 1]).FontSize(12);
                    cell_paragraph.Alignment = Alignment.center;
                }

                int r = 0;

                for (int row = 0; row < docxTableStudent.Count; row++)
                {
                    while (docxTableGroup[row] == docxTableGroup[r] && docxTableLesson[row] == docxTableLesson[r] && r < docxTableStudent.Count)
                    {
                        r++;
                        if (r == docxTableStudent.Count)
                        {
                            break;
                        }
                    }
                    if (r - row > 1)
                    {
                        table.MergeCellsInColumn(0, row + 1, r);
                        table.MergeCellsInColumn(2, row + 1, r);
                        row = r - 1;
                    }
                }
            }
            else
            {
                for (int row = 1; row < docxTableStudent.Count + 1; row++)
                {
                    table.Rows[row].Cells[0].RemoveParagraphAt(0);
                    table.Rows[row].Cells[1].RemoveParagraphAt(0);
                    table.Rows[row].Cells[2].RemoveParagraphAt(0);
                    Paragraph cell_paragraph = table.Rows[row].Cells[0].InsertParagraph(docxTableGroup[row - 1]).FontSize(12);
                    cell_paragraph.Alignment = Alignment.center;

                    cell_paragraph = table.Rows[row].Cells[1].InsertParagraph(docxTableStudent[row - 1]).FontSize(12);
                    cell_paragraph.Alignment = Alignment.center;

                    cell_paragraph = table.Rows[row].Cells[2].InsertParagraph(docxTableLesson[row - 1]).FontSize(12);
                    cell_paragraph.Alignment = Alignment.center;
                }
            }

            table.SetColumnWidth(0, 70);
            table.SetColumnWidth(1, 200);
            table.SetColumnWidth(2, 200);
            table.Alignment = Alignment.left;
            document.Save();
        }


        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if ((bool)dataGridView1[0, i].Value)
                {
                    docxTableGroup.Add(dataGridView1[1, i].Value.ToString());
                    docxTableStudent.Add(dataGridView1[2, i].Value.ToString());
                    docxTableLesson.Add(dataGridView1[5, i].Value.ToString() + "\n(" + dataGridView1[7, i].Value.ToString() + ")");
                }
            }
            for (int i = 0; i < docxTableStudent.Count; i++)
            {
                for (int j = 0; j < docxTableStudent.Count; j++)
                {
                    if (docxTableStudent[i] == docxTableStudent[j] && i != j)
                    {
                        MessageBox.Show(docxTableStudent[i] + " не может присутствовать \nна двух пересдачах одновременно.", "Ошибка");
                        docxTableStudent.Clear();
                        docxTableLesson.Clear();
                        docxTableGroup.Clear();
                        return;
                    }
                }
            }
            dataGridView1.Rows.Clear();

            checkedListBox1.Items.Clear();

            DocX document = null;
            if (documentPut.Contains(".docx"))
            {
                document = DocX.Load(documentPut);
            }
            else
            {
                document = DocX.Load(documentPut + ".docx");
            }
            document.InsertParagraph("").Font("Times New Roman").FontSize(14).Alignment = Xceed.Document.NET.Alignment.left;
            document.InsertParagraph(comboBox2.Text + ": " + teachers + ".").Font("Times New Roman").FontSize(14).Alignment = Xceed.Document.NET.Alignment.left;
            document.InsertParagraph("Дата и время: ").Font("Times New Roman").FontSize(14).Append(dateTimePicker1.Text + ", " + textBox2.Text).Bold().Font("Times New Roman").FontSize(14).Alignment = Xceed.Document.NET.Alignment.left;

            document.Save();
            DocxTableCreator(sender, e);
            docxTableStudent.Clear();
            docxTableLesson.Clear();
            docxTableGroup.Clear();
        }


        async void ComboInstitutes()
        {
            SqlCommand commandSelect = new SqlCommand($"SELECT [full_name_department] FROM [departments]", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await commandSelect.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    comboBox1.Items.Add(Convert.ToString(sqlReader["full_name_department"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            Program.sqlConnection.Close();
        }



        async void ChekedListLesson(string name_department)
        {
            SqlCommand commandSelect = new SqlCommand($"SELECT [lesson] FROM [arrears] WHERE [departament] = N'{name_department}'", Program.sqlConnection);
            commandSelect.Parameters.AddWithValue("departament", name_department);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await commandSelect.ExecuteReaderAsync();
                List<string> zapolnenie = new List<string> { };
                while (await sqlReader.ReadAsync())
                {
                    if (!zapolnenie.Contains(Convert.ToString(sqlReader["lesson"])))
                    {
                        zapolnenie.Add(Convert.ToString(sqlReader["lesson"]));
                    }
                }
                zapolnenie.Sort();
                for (int i = 0; i < zapolnenie.Count; i++)
                {
                    checkedListBox1.Items.Add(zapolnenie[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            Program.sqlConnection.Close();
        }


        private async void button6_Click(object sender, EventArgs e)
        {
            vremyaPeresdachi.Add(dateTimePicker1.Value);
            auditoriyaPeresdachi.Add(textBox2.Text);
            teachers_docx();
            Program.Connect();
            await Program.sqlConnection.OpenAsync();
            string full_name_department = comboBox1.Text;
            ChekedListLesson(NameDepartmentSelect(full_name_department));
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && textBox2.Text != "")
                button6.Enabled = true;
            else
            {
                button6.Enabled = false;
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.CheckedIndices.Contains(i))
                {
                    DataGridInfo(checkedListBox1.Items[i].ToString());
                }
            }
        }

        void DataGridInfo(string VariableOfChangeLesson)
        {
            Program.Connect();
            Program.sqlConnection.Open();
            SqlCommand commandSelect = new SqlCommand($"SELECT * FROM [arrears] WHERE [lesson] = N'{VariableOfChangeLesson}'", Program.sqlConnection);
            commandSelect.Parameters.AddWithValue("lesson", VariableOfChangeLesson);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = commandSelect.ExecuteReader();
                while (sqlReader.Read())
                {
                    dataGridView1.Rows.Add(false, Convert.ToString(sqlReader["name_groups"]), Convert.ToString(sqlReader["name_students"]), Convert.ToString(sqlReader["numberOfSemestr"]), Convert.ToString(sqlReader["semiannual"]), Convert.ToString(sqlReader["lesson"]), Convert.ToString(sqlReader["departament"]), Convert.ToString(sqlReader["typeOfControl"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            Program.sqlConnection.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
                dataGridView1[e.ColumnIndex, e.RowIndex].Value = !(bool)dataGridView1[e.ColumnIndex, e.RowIndex].Value;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkedListBox2.Items.Clear();
            string full_name_department = comboBox1.Text;

            Program.sqlConnection.Open();
            FillingComboTeachers(IdDepartmentSelect(full_name_department));
            if (comboBox1.Text != "")
            {
                button3.Enabled = true;
            }
            else
            {
                button3.Enabled = false;
            }
        }

        private void teachers_docx()
        {
            teachers = "";
            List<string> teachers_list = new List<string> { };
            
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (checkedListBox2.GetItemChecked(i))
                {
                    teachers_list.Add(checkedListBox2.Items[i].ToString());
                }
            }
            for (int i = 0; i < teachers_list.Count; i++)
            {
                teachers += teachers_list[i];
                if (i != teachers_list.Count - 1)
                {
                    teachers += ", ";
                }
            }
        }
    }
}
