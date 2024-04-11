using SpreadsheetLight;
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

namespace Bitacorafinal
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            ConfigurarComboBox();
            CargarComboBox();
            CargarComboBox2();
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        private void ConfigurarComboBox()
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = true;
            panel4.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("insert into geo(asignatura,nprofe,hentrada,hsalida) values ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox7.Text + "','" + textBox8.Text + "')", cn);
                cn.Open();
                cmd.ExecuteNonQuery();

            }

            CargarComboBox();
            LimpiarInterfazre();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccione una materia antes de agregar los datos");
                return;
            }


            string query = "INSERT INTO georegis (idMateria,NumAlumnos, NumClase, hclase,eqUtilizado) values (" + comboBox2.SelectedValue + "," + numericUpDown1.Value + "," + numericUpDown3.Value + "," + numericUpDown4.Value + "," + numericUpDown2.Value + ")";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlCommand cmd = new SqlCommand(query, cn);
                cn.Open();
                cmd.ExecuteNonQuery();

            }

            LimpiarInterfazma();


        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccione una materia antes de agregar los datos");
                return;
            }


            string query = "UPDATE georegis SET NumAlumnos = " + numericUpDown1.Value + " , hclase = " + numericUpDown4.Value + ", eqUtilizado = " + numericUpDown2.Value + " WHERE NumClase =  " + numericUpDown3.Value + " and idMateria = " + comboBox2.SelectedValue + "";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlCommand cmd = new SqlCommand(query, cn);
                cn.Open();
                cmd.ExecuteNonQuery();

            }

            LimpiarInterfazma();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccione una materia antes de agregar los datos");
                return;
            }


            string query = "SELECT NumClase AS 'Número de clase', hentrada AS 'hora de entrada',hsalida AS 'hora de salida',NumAlumnos AS 'Numero de Alumnos',hclase AS 'Horas de clase',eqUtilizado AS 'Equipos Utilizados' FROM geo,georegis WHERE geo.idMateria = " + comboBox3.SelectedValue + " and georegis.idMateria = " + comboBox3.SelectedValue + "";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlDataAdapter da = new SqlDataAdapter(query, cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                cn.Open();


            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccione una materia antes de eliminar datos");
                return;
            }


            string query = "DELETE FROM georegis WHERE idMateria = " + comboBox3.SelectedValue + " ";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlDataAdapter da = new SqlDataAdapter(query, cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                cn.Open();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string query = "DELETE FROM georegis";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlDataAdapter da = new SqlDataAdapter(query, cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                cn.Open();
            }

        }

        private void CargarComboBox()
        {
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM geo";
                SqlCommand cmd = new SqlCommand(query, cn);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable tablaMaterias = new DataTable();
                    adapter.Fill(tablaMaterias);
                    comboBox2.DisplayMember = "IdAsignatura";
                    comboBox2.ValueMember = "idMateria";
                    comboBox2.DataSource = tablaMaterias;
                    comboBox3.DisplayMember = "IdAsignatura";
                    comboBox3.ValueMember = "idMateria";
                    comboBox3.DataSource = tablaMaterias;
                }
                cn.Open();
                cmd.ExecuteNonQuery();
            }



        }



        private void LimpiarInterfazre()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";

        }

        private void LimpiarInterfazma()
        {

            comboBox2.SelectedValue = -1;
            numericUpDown1.Value = 0;
            numericUpDown3.Value = 0;
            numericUpDown4.Value = 0;
            numericUpDown2.Value = 0;


        }


        private void CargarComboBox2()
        {
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM geo";
                SqlCommand cmd = new SqlCommand(query, cn);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable tablaMaterias = new DataTable();
                    adapter.Fill(tablaMaterias);
                    comboBox7.DisplayMember = "IdAsignatura";
                    comboBox7.ValueMember = "idMateria";
                    comboBox7.DataSource = tablaMaterias;


                }
                cn.Open();
                cmd.ExecuteNonQuery();
            }

            using (SqlConnection cnn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM geo";
                SqlCommand cmd = new SqlCommand(query, cnn);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable tablaMaterias1 = new DataTable();
                    adapter.Fill(tablaMaterias1);
                    comboBox8.DisplayMember = "IdAsignatura";
                    comboBox8.ValueMember = "idMateria";
                    comboBox8.DataSource = tablaMaterias1;


                }
                cnn.Open();
                cmd.ExecuteNonQuery();
            }

            using (SqlConnection cnnn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM geo";
                SqlCommand cmd = new SqlCommand(query, cnnn);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable tablaMaterias2 = new DataTable();
                    adapter.Fill(tablaMaterias2);
                    comboBox1.DisplayMember = "IdAsignatura";
                    comboBox1.ValueMember = "idMateria";
                    comboBox1.DataSource = tablaMaterias2;


                }
                cnnn.Open();
                cmd.ExecuteNonQuery();
            }

            using (SqlConnection cnnnn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM geo";
                SqlCommand cmd = new SqlCommand(query, cnnnn);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable tablaMaterias3 = new DataTable();
                    adapter.Fill(tablaMaterias3);
                    comboBox4.DisplayMember = "IdAsignatura";
                    comboBox4.ValueMember = "idMateria";
                    comboBox4.DataSource = tablaMaterias3;


                }
                cnnnn.Open();
                cmd.ExecuteNonQuery();
            }

            using (SqlConnection cn5 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM geo";
                SqlCommand cmd = new SqlCommand(query, cn5);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable tablaMaterias4 = new DataTable();
                    adapter.Fill(tablaMaterias4);
                    comboBox5.DisplayMember = "IdAsignatura";
                    comboBox5.ValueMember = "idMateria";
                    comboBox5.DataSource = tablaMaterias4;


                }
                cn5.Open();
                cmd.ExecuteNonQuery();
            }

            using (SqlConnection cn6 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM geo";
                SqlCommand cmd = new SqlCommand(query, cn6);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable tablaMaterias5 = new DataTable();
                    adapter.Fill(tablaMaterias5);
                    comboBox6.DisplayMember = "IdAsignatura";
                    comboBox6.ValueMember = "idMateria";
                    comboBox6.DataSource = tablaMaterias5;


                }
                cn6.Open();
                cmd.ExecuteNonQuery();
            }



        }













        private void button1_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;

            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("SELECT COUNT(asignatura) FROM geo ", cn);
                cn.Open();
                cmd.ExecuteNonQuery();

                int count = (int)cmd.ExecuteScalar();

                label20.Text = count.ToString();
                HorasTotales.Total6 += count;
                HorasTotales.Total12 = count;

            }

            using (SqlConnection cn2 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("SELECT SUM(MaxAlumnos) FROM (SELECT MAX(NumAlumnos) AS MaxAlumnos FROM georegis GROUP BY idMateria) AS Subconsulta", cn2);
                cn2.Open();
                cmd.ExecuteNonQuery();

                decimal sum = Convert.ToDecimal(cmd.ExecuteScalar());

                label19.Text = sum.ToString();
                HorasTotales.Total5 += sum;
                HorasTotales.Total11 = sum;

            }

            using (SqlConnection cn3 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("SELECT sum(hclase) from georegis", cn3);
                cn3.Open();
                cmd.ExecuteNonQuery();

                decimal sum = Convert.ToDecimal(cmd.ExecuteScalar());

                label18.Text = sum.ToString();
                HorasTotales.Total4 += sum;
                HorasTotales.Total10 = sum;


            }




        }

        private void button12_Click(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            using (SqlConnection cn5 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                try
                {
                    cn5.Open();

                    string query = "SELECT SUM(hclase) FROM civregis INNER JOIN civil ON civregis.idMateria = civil.idMateria WHERE civil.idMateria BETWEEN @materiaInicio AND @materiaFin";

                    using (SqlCommand cmd = new SqlCommand(query, cn5))
                    {

                        int materiaInicio = Convert.ToInt32(comboBox1.SelectedValue);
                        int materiaFin = Convert.ToInt32(comboBox4.SelectedValue);


                        cmd.Parameters.AddWithValue("@materiaInicio", materiaInicio);
                        cmd.Parameters.AddWithValue("@materiaFin", materiaFin);


                        object result = cmd.ExecuteScalar();


                        if (result != DBNull.Value && result != null)
                        {
                            decimal sum = Convert.ToDecimal(result);
                            label22.Text = sum.ToString();
                            HorasTotales.Total2 += sum;
                            HorasTotales.Total8 = sum;
                        }
                        else
                        {

                            label22.Text = "0";
                        }
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            using (SqlConnection cn4 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                try
                {
                    cn4.Open();

                    string query = "SELECT SUM(hclase) FROM civregis INNER JOIN civil ON civregis.idMateria = civil.idMateria WHERE civil.idMateria BETWEEN @materiaInicio AND @materiaFin";

                    using (SqlCommand cmd = new SqlCommand(query, cn4))
                    {

                        int materiaInicio = Convert.ToInt32(comboBox7.SelectedValue);
                        int materiaFin = Convert.ToInt32(comboBox8.SelectedValue);


                        cmd.Parameters.AddWithValue("@materiaInicio", materiaInicio);
                        cmd.Parameters.AddWithValue("@materiaFin", materiaFin);


                        object result = cmd.ExecuteScalar();


                        if (result != DBNull.Value && result != null)
                        {
                            decimal sum = Convert.ToDecimal(result);
                            label21.Text = sum.ToString();
                            HorasTotales.Total += sum;
                            HorasTotales.Total7 = sum;
                        }
                        else
                        {

                            label21.Text = "0";
                        }
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            using (SqlConnection cn4 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                try
                {
                    cn4.Open();

                    string query = "SELECT SUM(hclase) FROM civregis INNER JOIN civil ON civregis.idMateria = civil.idMateria WHERE civil.idMateria BETWEEN @materiaInicio AND @materiaFin";

                    using (SqlCommand cmd = new SqlCommand(query, cn4))
                    {

                        int materiaInicio = Convert.ToInt32(comboBox5.SelectedValue);
                        int materiaFin = Convert.ToInt32(comboBox6.SelectedValue);


                        cmd.Parameters.AddWithValue("@materiaInicio", materiaInicio);
                        cmd.Parameters.AddWithValue("@materiaFin", materiaFin);


                        object result = cmd.ExecuteScalar();


                        if (result != DBNull.Value && result != null)
                        {
                            decimal sum = Convert.ToDecimal(result);
                            label23.Text = sum.ToString();
                            HorasTotales.Total3 += sum;
                            HorasTotales.Total9 = sum;
                        }
                        else
                        {

                            label23.Text = "0";
                        }
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = true;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument();

            int IC = 1;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                sl.SetCellValue(1, IC, column.HeaderText.ToString());
                IC++;
                sl.AutoFitColumn(1, IC);

            }

            int IR = 2;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                sl.SetCellValue(IR, 1, row.Cells[0].Value.ToString());
                sl.SetCellValue(IR, 2, row.Cells[1].Value.ToString());
                sl.SetCellValue(IR, 3, row.Cells[2].Value.ToString());
                sl.SetCellValue(IR, 4, row.Cells[3].Value.ToString());
                sl.SetCellValue(IR, 5, row.Cells[4].Value.ToString());
                sl.SetCellValue(IR, 6, row.Cells[5].Value.ToString());
                IR++;
            }

            // Crear un nuevo diálogo de guardar archivo
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Guardar datos de la materia";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)

            {
                string filePath = saveFileDialog.FileName;
                sl.SaveAs(filePath);
                MessageBox.Show("Datos guardados exitosamente", "Datos Guardados", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
        }
    }
}
