using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using ClosedXML.Excel;


namespace project1_Parser
{
    public partial class Form1 : Form
    {
        int countOfRow = 1;
       
        public Form1()
        {
            Program.f1 = this;
            InitializeComponent();
            UpdatetableOpen();
            dataGridView2.ColumnHeadersVisible = false;
            dataGridView3.ColumnHeadersVisible = false;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void ФайлToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        static IXLWorksheet _WorkSheet;

        

        async void UpdateTable()
        {
            SqlCommand commandSelect = new SqlCommand("SELECT * FROM [arrears]", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                dataGridView1.Rows.Clear();
                sqlReader = await commandSelect.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    dataGridView1.Rows.Add(Convert.ToString(sqlReader["name_groups"]), Convert.ToString(sqlReader["name_students"]), Convert.ToString(sqlReader["numberOfSemestr"]),  Convert.ToString(sqlReader["semiannual"]), Convert.ToString(sqlReader["lesson"]), Convert.ToString(sqlReader["departament"]), Convert.ToString(sqlReader["typeOfControl"]));
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
                Program.connected = false;
            }
        }

        async void DeleteFromTable()
        {
            SqlCommand commandSelect = new SqlCommand("DELETE FROM arrears;", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await commandSelect.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    await commandSelect.ExecuteNonQueryAsync();
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
                Program.connected = false;
            }
        }

        async void ChangeLesson(string VariableOfChangeLesson)
        {
            SqlCommand commandSelect = new SqlCommand($"SELECT lesson FROM [arrears] WHERE [name_students] = N'{VariableOfChangeLesson}'", Program.sqlConnection);
            commandSelect.Parameters.AddWithValue("name_students", VariableOfChangeLesson);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await commandSelect.ExecuteReaderAsync();
                
                while (await sqlReader.ReadAsync())
                {
                    comboBox2.Items.Add(Convert.ToString(sqlReader["lesson"]));
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
                Program.connected = false;
            }
        }

        async void DeleteArrear(string name_student, string lesson)
        {
            SqlCommand commandSelect = new SqlCommand($"DELETE FROM [arrears] WHERE [name_students] = N'{name_student}' AND [lesson] = N'{lesson}'", Program.sqlConnection);
            commandSelect.Parameters.AddWithValue("name_students", name_student);
            commandSelect.Parameters.AddWithValue("lesson", lesson);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await commandSelect.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    await commandSelect.ExecuteNonQueryAsync();
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
                Program.connected = false;
            }
        }

        async void AddArrear(string name_group, string name_student, string numberOfSemestr, string semiannual, string lesson, string departament, string typeOfControl)
        {
            Program.Connect();
            await Program.sqlConnection.OpenAsync();
            SqlCommand commandSelect = new SqlCommand($"INSERT INTO [arrears]  VALUES( N'{name_group}', N'{name_student}', N'{numberOfSemestr}', N'{semiannual}', N'{lesson}', N'{departament}', N'{typeOfControl}')", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await commandSelect.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    await commandSelect.ExecuteNonQueryAsync();
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
                Program.connected = false;
            }
        }

        void DeleteDepartment(string name_department)
        {
            SqlCommand commandSelect = new SqlCommand($"DELETE FROM [departments] WHERE [name_department] = N'{name_department}'", Program.sqlConnection);
            commandSelect.Parameters.AddWithValue("name_department", name_department);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = commandSelect.ExecuteReader();

                while (sqlReader.Read())
                {
                    commandSelect.ExecuteNonQueryAsync();
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
                Program.connected = false;
            }
        }

        void DepartmentAdd(string name_department, string full_name_department)
        {
            SqlCommand command = new SqlCommand($"INSERT [departments]([name_department], [full_name_department]) VALUES (N'{name_department}', N'{full_name_department}')", Program.sqlConnection);
            command.Parameters.AddWithValue("name_department", name_department);
            command.Parameters.AddWithValue("full_name_department", full_name_department);
            try
            {
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Program.sqlConnection.Close();
                Program.connected = false;
            }
        }

        string IdDepartmentSelect(string name_department)
        {
            string id_department;
            SqlCommand command1 = new SqlCommand($"SELECT [id_department] FROM [departments] WHERE [name_department] = N'{name_department}'", Program.sqlConnection);
            command1.Parameters.AddWithValue("name_department", name_department);
            id_department = Convert.ToString(command1.ExecuteScalar());
            return id_department;
        }

        void TeacherAdd(string id_department, string name_teacher)
        {
            SqlCommand command2 = new SqlCommand($"INSERT [teachers]([name_teachers], [id_department]) VALUES (N'{name_teacher}', N'{id_department}')", Program.sqlConnection);
            command2.Parameters.AddWithValue("name_teacher", name_teacher);
            command2.Parameters.AddWithValue("id_department", id_department);

            try
            {
                command2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Program.sqlConnection.Close();
                Program.connected = false;
            }
        }

        void TeacherDelete(string id_department, string name_teacher)
        {
            SqlCommand command2 = new SqlCommand($"DELETE FROM [teachers] WHERE [name_teachers] = N'{name_teacher}' AND [id_department] = N'{id_department}'", Program.sqlConnection);
            command2.Parameters.AddWithValue("name_teacher", name_teacher);
            command2.Parameters.AddWithValue("id_department", id_department);

            try
            {
                command2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Program.sqlConnection.Close();
                Program.connected = false;
            }
        }

        void TeacherDeleteWithDepartment(string id_department)
        {
            SqlCommand command1 = new SqlCommand($"DELETE FROM [teachers] WHERE [id_department] = N'{id_department}'", Program.sqlConnection);
            command1.Parameters.AddWithValue("id_department", id_department);

            try
            {
                command1.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Program.sqlConnection.Close();
                Program.connected = false;
            }
        }

        void FillingTableDepartment()
        {
            SqlCommand commandSelect = new SqlCommand($"SELECT [name_department] FROM [departments] ORDER BY [name_department]", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = commandSelect.ExecuteReader();

                while (sqlReader.Read())
                {
                    dataGridView2.Rows.Add(Convert.ToString(sqlReader["name_department"]));
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
        }

        void FillingTableTeachers()
        {
            SqlCommand commandSelect = new SqlCommand($"SELECT [name_teachers] FROM [teachers] ORDER BY [name_teachers]", Program.sqlConnection);
            SqlDataReader sqlReader1 = null;
            try
            {
                sqlReader1 =  commandSelect.ExecuteReader();

                while ( sqlReader1.Read())
                {
                    dataGridView3.Rows.Add(Convert.ToString(sqlReader1["name_teachers"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader1 != null)
                    sqlReader1.Close();
                Program.sqlConnection.Close();
                Program.connected = false;
            }
        }

        void FillingComboDepartmens()
        {
            SqlCommand commandSelect = new SqlCommand($"SELECT [name_department] FROM [departments] ORDER BY [name_department]", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = commandSelect.ExecuteReader();

                while (sqlReader.Read())
                {
                   comboBox3.Items.Add(Convert.ToString(sqlReader["name_department"]));
                   comboBox4.Items.Add(Convert.ToString(sqlReader["name_department"]));
                   comboBox5.Items.Add(Convert.ToString(sqlReader["name_department"]));
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
                Program.connected = false;
            }
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
                    comboBox6.Items.Add(Convert.ToString(sqlReader["name_teachers"]));
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
                Program.connected = false;
            }
        }

        private async void UpdatetableOpen()
        {
            Program.Connect();
            await Program.sqlConnection.OpenAsync();

            SqlCommand commandSelect = new SqlCommand("SELECT * FROM [arrears]", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                dataGridView1.Rows.Clear();
                sqlReader = await commandSelect.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    dataGridView1.Rows.Add(Convert.ToString(sqlReader["name_groups"]), Convert.ToString(sqlReader["name_students"]), Convert.ToString(sqlReader["numberOfSemestr"]), Convert.ToString(sqlReader["semiannual"]), Convert.ToString(sqlReader["lesson"]), Convert.ToString(sqlReader["departament"]), Convert.ToString(sqlReader["typeOfControl"]));
                    string proverka = Convert.ToString(sqlReader["name_students"]);
                    if (comboBox1.Items.Count == 0)
                    {
                        comboBox1.Items.Add(proverka);
                    }else if (proverka != comboBox1.Items[comboBox1.Items.Count - 1].ToString() && proverka != "")
                    {
                        comboBox1.Items.Add(Convert.ToString(sqlReader["name_students"]));
                    }
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
                Program.connected = false;
            }
        }

        private  void UpdateComboboxOpen()
        {
            Program.sqlConnection.Open();

            SqlCommand commandSelect = new SqlCommand("SELECT * FROM [arrears]", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                dataGridView1.Rows.Clear();
                sqlReader = commandSelect.ExecuteReader();
                while (sqlReader.Read())
                {
                   string proverka = Convert.ToString(sqlReader["name_students"]);
                    if (comboBox1.Items.Count == 0)
                    {
                        comboBox1.Items.Add(proverka);
                    }
                    else if (proverka != comboBox1.Items[comboBox1.Items.Count - 1].ToString() && proverka != "")
                    {
                        comboBox1.Items.Add(Convert.ToString(sqlReader["name_students"]));
                    }

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
                Program.connected = false;
            }
        }

        private async void ОткрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var workbook = new XLWorkbook(openFileDialog1.FileName);
                _WorkSheet = workbook.Worksheet(1);

                Program.Connect();
                await Program.sqlConnection.OpenAsync();
                
                while (!_WorkSheet.Cell(countOfRow , 1).IsEmpty())
                {
                    countOfRow++;

                    SqlCommand command = new SqlCommand("INSERT INTO [arrears] ([name_groups], [name_students], [numberOfSemestr], [semiannual], [lesson], [departament], [typeOfControl]) VALUES(@name_groups, @name_students, @numberOfSemestr, @semiannual, @lesson, @departament, @typeOfControl)", Program.sqlConnection);
                    
                    command.Parameters.AddWithValue("name_groups",_WorkSheet.Cell(countOfRow, 1).Value.ToString().TrimEnd());
                    command.Parameters.AddWithValue("name_students", _WorkSheet.Cell(countOfRow, 2).Value.ToString().TrimEnd());
                    command.Parameters.AddWithValue("numberOfSemestr", _WorkSheet.Cell(countOfRow, 3).Value.ToString().TrimEnd());
                    command.Parameters.AddWithValue("semiannual", _WorkSheet.Cell(countOfRow, 4).Value.ToString().TrimEnd());
                    command.Parameters.AddWithValue("lesson", _WorkSheet.Cell(countOfRow, 5).Value.ToString().TrimEnd());
                    command.Parameters.AddWithValue("departament", _WorkSheet.Cell(countOfRow, 6).Value.ToString().TrimEnd());
                    command.Parameters.AddWithValue("typeOfControl", _WorkSheet.Cell(countOfRow, 7).Value.ToString().TrimEnd());

                    string proverka = _WorkSheet.Cell(countOfRow, 2).Value.ToString(); 
                    if (proverka != _WorkSheet.Cell(countOfRow + 1, 2).Value.ToString() && !_WorkSheet.Cell(countOfRow, 1).IsEmpty())
                        comboBox1.Items.Add(_WorkSheet.Cell(countOfRow, 2).Value.ToString().TrimEnd());
                    if (!_WorkSheet.Cell(countOfRow, 1).IsEmpty())
                    {
                        await command.ExecuteNonQueryAsync();
                    }
                }
                UpdateTable();
            }
        }
        
        private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Program.sqlConnection != null && Program.sqlConnection.State != ConnectionState.Closed)
                Program.sqlConnection.Close();
        }

        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.sqlConnection != null && Program.sqlConnection.State != ConnectionState.Closed)
                Program.sqlConnection.Close();
            this.Close();
        }

        private async void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Text = "";
            comboBox2.Items.Clear();
            string VariableOfChangeLesson = comboBox1.Text;
            await Program.sqlConnection.OpenAsync();
            ChangeLesson(VariableOfChangeLesson);
        }

        private async void Button2_Click(object sender, EventArgs e)
        {
            string name_student = comboBox1.Text;
            string lesson = comboBox2.Text;
            await Program.sqlConnection.OpenAsync();
            DeleteArrear(name_student, lesson);
            comboBox1.Text = "";
            comboBox2.Text = "";
        }

        private void ОбновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Program.sqlConnection.OpenAsync();
            UpdateTable();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Owner = this;
            f2.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox1.Text == "Полное наименование кафедры!" || comboBox3.Items.Contains(textBox1.Text) || textBox1.Text == "Успешно!"
                || textBox10.Text == "" || textBox1.Text == "Абревиатура")
            {
                return;
            }
            string name_department = textBox10.Text;
            string full_name_department = textBox1.Text;
            Program.sqlConnection.Open();
            DepartmentAdd(name_department, full_name_department);
            tabPage5_Enter(sender, e);
            textBox1.Text = "Успешно!";
            textBox10.Text = "Успешно!";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox3.Text == "")
            {
                MessageBox.Show("Выберите кафедру.", "Предупреждение");
                return;
            }
            if (textBox2.Text == "Фамилия И.О.")
            {
                MessageBox.Show("Введите ФИО преподавателя в указанном формате", "Предупреждение");
                return;
            }
            string name_department = comboBox3.Text;
            string name_teacher = textBox2.Text;
            Program.sqlConnection.Open();
            TeacherAdd(IdDepartmentSelect(name_department), name_teacher);
            textBox2.Text = "Успешно!";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "В базе данных могут храниться данные связаные с выбранной кафедрой.\n" +
                "Они будут удалены вместе с кафедрой.\n" +
                "\nПродолжить?",
                "Предупреждение",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                string name_department = comboBox4.Text;
                Program.sqlConnection.Open();
                TeacherDeleteWithDepartment(IdDepartmentSelect(name_department));
                Program.sqlConnection.Open();
                DeleteDepartment(name_department);
                tabPage5_Enter(sender, e);
                comboBox4.Text = "";
            }
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox5.Text == "" || comboBox6.Text == "")
                return;
            string name_department = comboBox5.Text;
            string name_teacher = comboBox6.Text;
            Program.sqlConnection.Open();
            TeacherDelete(IdDepartmentSelect(name_department), name_teacher);
            comboBox5.Text = "";
            comboBox6.Text = "";
            comboBox5_SelectedIndexChanged(sender, e);

        }

        private void tabPage6_Enter(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            Program.sqlConnection.Open();
            FillingTableDepartment();
            FillingTableTeachers();
        }

        private void tabPage5_Enter(object sender, EventArgs e)
        {
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();
            comboBox6.Items.Clear();
            Program.sqlConnection.Open();
            FillingComboDepartmens();
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "Фамилия И.О." || textBox2.Text == "Успешно!")
            {
                textBox2.Text = "";
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "Полное наименование кафедры!" || textBox1.Text == "Успешно!")
            {
                textBox1.Text = "";
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox6.Items.Clear();
            string name_department = comboBox5.Text;
            
            Program.sqlConnection.Open();
            FillingComboTeachers(IdDepartmentSelect(name_department));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "" || textBox4.Text != "" || textBox5.Text != "" || textBox6.Text != "" 
                || textBox7.Text != "" || textBox8.Text != "" || textBox9.Text != "")
            {
                AddArrear(textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text);
            }
        }

        private async void tabPage1_Enter(object sender, EventArgs e)
        {
            if (!Program.connected) 
            {
                Program.Connect();
                await Program.sqlConnection.OpenAsync();
                UpdateTable();
            } 
        }

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            UpdateComboboxOpen();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            Form3 f3 = form3;
            f3.Owner = this;
            f3.Show();
        }

        private void textBox10_Click(object sender, EventArgs e)
        {
            if (textBox10.Text == "Абревиатура" || textBox1.Text == "Успешно!")
            {
                textBox10.Text = "";
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                видToolStripMenuItem.Visible = true;
            }
            else
            {
                видToolStripMenuItem.Visible = false;
            }
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Program.sqlConnection.OpenAsync();
            DeleteFromTable();
            Program.sqlConnection.OpenAsync();
            UpdateTable();
        }
    }
}
