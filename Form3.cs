using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace project1_Parser
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            string[] combobox1Items = { "Группы", "Дисциплины", "Кафедры" };
            string[] combobox2Items = { "1", "2", "3", "4", "5", "6", "7", "8" };
            comboBox1.Items.AddRange(combobox1Items);
            comboBox2.Items.AddRange(combobox2Items);
        }
        Dictionary<string, double> Data = new Dictionary<string, double>();
        Dictionary<string, double> Data1 = new Dictionary<string, double>();
        private void Form3_Load(object sender, EventArgs e)
        {
            

        }

        private void UpdateComboboxLoad()
        {
            Program.Connect();
            Program.sqlConnection.Open();
            SqlCommand commandSelect = new SqlCommand("SELECT * FROM [arrears]", Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {
                chart1.ChartAreas.Clear();
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
            }
        }

        void ChekedListFill(int comboBoxIndex, string numberOFSemestr)
        {
            string comandText = "";
            string temp = "";
            switch (comboBoxIndex)
            {
                case 0:
                    comandText = $"SELECT [name_groups] FROM [arrears] WHERE [numberOFSemestr] = N'{numberOFSemestr}'";
                    temp = "name_groups";
                    break;
                case 1:
                    comandText = $"SELECT [lesson] FROM [arrears] WHERE [numberOFSemestr] = N'{numberOFSemestr}'";
                    temp = "lesson";
                    break;
                case 2:
                    comandText = $"SELECT [departament] FROM [arrears] WHERE [numberOFSemestr] = N'{numberOFSemestr}'";
                    temp = "departament";
                    break;
            }

            Program.Connect();
            Program.sqlConnection.Open();
            SqlCommand commandSelect = new SqlCommand(comandText, Program.sqlConnection);
            commandSelect.Parameters.AddWithValue("numberOFSemestr", numberOFSemestr);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = commandSelect.ExecuteReader();
                List<string> zapolnenie = new List<string> { };
                while (sqlReader.Read())
                {
                    if (!zapolnenie.Contains(Convert.ToString(sqlReader[temp])))
                    {
                        zapolnenie.Add(Convert.ToString(sqlReader[temp]));
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox2.Text != "")
            {
                checkedListBox1.Items.Clear();
                ChekedListFill(comboBox1.SelectedIndex, comboBox2.Text);
            }
            chart1.ChartAreas[0].AxisX.Title = comboBox1.Text;
            chart1.ChartAreas[0].AxisY.Title = "Кол-во";
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox2.Text != "")
            {
                checkedListBox1.Items.Clear();
                ChekedListFill(comboBox1.SelectedIndex, comboBox2.Text);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Data.Clear();
            Data1.Clear();
            chart1.Series[0].Points.Clear();
            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                List<string> temp = new List<string> { };
                Data.Add(checkedListBox1.CheckedItems[i].ToString(), Countforchart(checkedListBox1.CheckedItems[i].ToString()));
                Data1.Add(checkedListBox1.CheckedItems[i].ToString() + "(" + (Countforchart1(checkedListBox1.CheckedItems[i].ToString(), temp) /
                     Countforchart(checkedListBox1.CheckedItems[i].ToString())*100.0).ToString("N2") + "%)", Countforchart1(checkedListBox1.CheckedItems[i].ToString(), temp));
                
            }

            for (int i = 0; i < Data.Count; i++)
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        this.chart1.Series[0].Points.AddXY(" ", Data.Values.ElementAt(i));
                        this.chart1.Series[0].Points.AddXY(Data1.Keys.ElementAt(i), Data1.Values.ElementAt(i));
                        break;
                    case 1:
                        this.chart1.Series[0].Points.AddXY(Data.Keys.ElementAt(i), Data.Values.ElementAt(i));
                        break;
                    case 2:
                        this.chart1.Series[0].Points.AddXY(Data.Keys.ElementAt(i), Data.Values.ElementAt(i));
                        break;
                }
                
            }
        }

        private double Countforchart(string str)
        {
            Program.Connect();
            Program.sqlConnection.Open();
            string commandstr = "";
            double cnt = 0;
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    commandstr = $"SELECT [name_groups] FROM [arrears] WHERE [name_groups] = N'{str}'";
                    break;
                case 1:
                    commandstr = $"SELECT * FROM [arrears] WHERE [lesson] = N'{str}'";
                    break;
                case 2:
                    commandstr = $"SELECT * FROM [arrears] WHERE [departament] = N'{str}'";
                    break;
            }

            SqlCommand commandSelect = new SqlCommand(commandstr, Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {

                sqlReader = commandSelect.ExecuteReader();
                while (sqlReader.Read())
                {
                    cnt += 1.0;
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
            return cnt;
        }

        private double Countforchart1(string str, List<string> temp)
        {
            Program.Connect();
            Program.sqlConnection.Open();
            string commandstr = "";
            temp.Clear();
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    commandstr = $"SELECT [name_students] FROM [arrears] WHERE [name_groups] = N'{str}'";
                    break;
                case 1:
                    commandstr = $"SELECT * FROM [arrears] WHERE [lesson] = N'{str}'";
                    break;
                case 2:
                    commandstr = $"SELECT * FROM [arrears] WHERE [departament] = N'{str}'";
                    break;
            }

            SqlCommand commandSelect = new SqlCommand(commandstr, Program.sqlConnection);
            SqlDataReader sqlReader = null;
            try
            {

                sqlReader = commandSelect.ExecuteReader();
                while (sqlReader.Read())
                {
                    string proverka = Convert.ToString(sqlReader["name_students"]);
                    if (temp.Count == 0)
                    {
                        temp.Add(proverka);
                    }
                    else if (proverka != temp[temp.Count - 1].ToString() && proverka != "")
                    {
                        temp.Add(Convert.ToString(sqlReader["name_students"]));
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
            }
            return temp.Count;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
