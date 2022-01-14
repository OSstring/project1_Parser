using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace project1_Parser
{
    public class sqlcomands
    {
        /*foreach(string variable in group)
                {
                    SqlCommand commandInsert = new SqlCommand("INSERT INTO [arrears] ([name_groups]) VALUES(@name_groups)", sqlConnection);
                    commandInsert.Parameters.AddWithValue("name_groups", variable);
                    await commandInsert.ExecuteNonQueryAsync();
                }*/
        //this.group.Add(_WorkSheet.Cell(countOfRow, 1).Value.ToString());
        //listBox1.Items.Add(_WorkSheet.Cell(countOfRow, 1).Value.ToString());
        //string a = _WorkSheet.Cell(countOfRow, 1).Value.ToString();
        //commandInsert.Parameters.AddWithValue("name_groups", a);
        //commandInsert.Parameters.Add(new SqlParameter("@name_groups", SqlDbType.VarChar) { Value = a });
        //commandInsert.Parameters.AddWithValue("@name_groups", _WorkSheet.Cell(countOfRow, 1).Value.ToString());
        //SqlCommand commandInsert = new SqlCommand($"INSERT INTO [arrears] ([name_groups]) VALUES({_WorkSheet.Cell(countOfRow, 1).Value.ToString()});", sqlConnection);

        //SqlCommand commandInsert = new SqlCommand("INSERT INTO [arrears] ([name_groups]) VALUES(@name_groups)", sqlConnection);


        //commandInsert.CommandType = CommandType.Text;



        /*string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\forgotten\source\repos\project1_Parser\Database1.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
            SqlDataReader sqlReader = null;
            SqlCommand commandSelect = new SqlCommand("SELECT * FROM [arrears]" + , sqlConnection);*/

        /*if (docxTableGroup[docxTableStudent.Count - 2] == docxTableGroup[docxTableStudent.Count - 1] && docxTableLesson[docxTableStudent.Count - 2] == docxTableLesson[docxTableStudent.Count - 1])
               {
                   int t = docxTableStudent.Count;
                   while (docxTableGroup[t - 2] == docxTableGroup[t - 1] && docxTableLesson[t - 2] == docxTableLesson[t - 1])
                   {
                       t--;
                   }
                   table.MergeCellsInColumn(0, t - 1, table.RowCount - 1);
                   table.MergeCellsInColumn(2, t - 1, table.RowCount - 1);
               }

              for (int row = 1; row < docxTableStudent.Count + 1; row++)
               {

                   Paragraph cell_paragraph = table.Rows[row].Cells[0].InsertParagraph(docxTableGroup[row-1]).FontSize(12);
                   cell_paragraph.Alignment = Alignment.center;

                   cell_paragraph = table.Rows[row].Cells[1].InsertParagraph(docxTableStudent[row - 1]).FontSize(12);
                   cell_paragraph.Alignment = Alignment.center;

                   cell_paragraph = table.Rows[row].Cells[2].InsertParagraph(docxTableLesson[row - 1]).FontSize(12);
                   cell_paragraph.Alignment = Alignment.center;
               }
               //ttemp = 1;
               //string proverka3 = table.Rows[2].Cells[0].ToString(), proverka4 = table.Rows[2].Cells[2].ToString();

               for (int row = 1; row < docxTableStudent.Count - 1; row++)
               {

                   if (table.Rows[row].Cells[0].ToString() == table.Rows[row+1].Cells[0].ToString() && table.Rows[row].Cells[2].ToString() == table.Rows[row + 1].Cells[2].ToString())
                   {
                       table.MergeCellsInColumn(0, row, row + 1);
                       table.MergeCellsInColumn(2, row, row + 1);
                   }
                   else
                   {
                       proverka3 = table.Rows[row + 1].Cells[0].ToString();
                       proverka4 = table.Rows[row + 1].Cells[2].ToString();
                   }

               }*/
        /*for (int i = 1; i < docxTableStudent.Count; i++)
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
                        ttemp++;
                    }
                }*//*for (int row = 1; row < docxTableStudent.Count; row++)
                {
                    while (table.Rows[row].Cells[0] == table.Rows[r].Cells[0] && table.Rows[row].Cells[2] == table.Rows[r].Cells[2] && r < docxTableStudent.Count)
                    {
                        r++;
                    }
                    if (r - row > 1)
                    {
                        table.MergeCellsInColumn(0, row, r);
                        table.MergeCellsInColumn(2, row, r);
                        r++;
                        row = r;
                    }
                }*/
    }
}
