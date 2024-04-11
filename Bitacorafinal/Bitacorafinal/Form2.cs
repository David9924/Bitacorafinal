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
using ClosedXML.Excel;
using OfficeOpenXml;
using LinqForEEPlus;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using System.IO;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Packaging;

namespace Bitacorafinal
{
    public partial class Form2 : Form
    {
        private int contador = 1;
        private int ultimoAnio = DateTime.Now.Year;
        public Form2()
        {
            //Se inicializan los datos
            InitializeComponent();
            ConfigurarComboBox();
            CargarComboBox();
            CargarComboBox2();
            //Se ocultan los paneles
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


        //Boton para ingresar o editar datos
        private void button3_Click(object sender, EventArgs e)
        {
            //Se muestran los paneles al presionar los botones
            panel1.Visible = true; 
            panel2.Visible = true;
            panel3.Visible = false;
            panel4.Visible = false;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string nuevoElemento = ObtenerNuevoElemento();


            MessageBox.Show($"Se agregó un nuevo elemento: {nuevoElemento}");


        }

        private string ObtenerNuevoElemento()
        {
            if (DateTime.Now.Year != ultimoAnio)
            {
                contador = 1;
                ultimoAnio = DateTime.Now.Year;
            }

            return contador++.ToString();

        }


        //Boton para ver estadisticas
        private void button4_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = true;
            panel4.Visible = false;

        }


        //Boton para agregar datos
        private void button6_Click(object sender, EventArgs e)
        {
            //Se hace una consulta de la base de datos creada
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                //Se hace una consulta agregando la asignatura, el nombre del profesor, la hora de entrada y salida en los textBox
                SqlCommand cmd = new SqlCommand("insert into civil(asignatura,nprofe,hentrada,hsalida) values ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox7.Text + "','" + textBox8.Text + "')", cn);
                cn.Open();
                cmd.ExecuteNonQuery();

            }

            CargarComboBox(); //Se cargan los datos de la base de datos
            LimpiarInterfazre(); //Se limpian los combobox


        }

        //Boton para editar una asignatura
        private void button7_Click(object sender, EventArgs e)
        {
            //Se coloca una condición de que no puedes ingresar datos si no seleccionaste alguna asignatura
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccione una materia antes de agregar los datos");
                    return;
            }

            //Se agregan los datos con INSERT desde la base de datos al agregar un valor en comboBox
            string query = "INSERT INTO civregis (idMateria,NumAlumnos, NumClase, hclase,eqUtilizado) values ("+comboBox2.SelectedValue+","+numericUpDown1.Value+","+numericUpDown3.Value+","+numericUpDown4.Value+","+numericUpDown2.Value+")";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlCommand cmd = new SqlCommand(query, cn);
                cn.Open();
                cmd.ExecuteNonQuery();

            }

            LimpiarInterfazma(); //Se limpia la interfaz


        }

        //Se cargan los datos en el programa
        private void CargarComboBox()
        {
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM civil"; //Se hace una consulta
                SqlCommand cmd = new SqlCommand(query, cn);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    //Se agregan los datos en los comboBox
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

        //Se cargan los datos para calcular las horas totales del laboratorio
        private void CargarComboBox2()
        {
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM civil";
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
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM civil";
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
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM civil";
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
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM civil";
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
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM civil";
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
                string query = "SELECT idMateria,CONCAT(idMateria, ' - ', asignatura, '-',nprofe) AS IdAsignatura FROM civil";
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


        //Función para limpiar los combobox y textbox cada que se agregan los datos.
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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
 
        }


        private void button9_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccione una materia antes de agregar los datos");
                return;
            }


            cargardatos();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Seleccione una materia antes de agregar los datos");
                return;
            }


            string query = "UPDATE civregis SET NumAlumnos = " + numericUpDown1.Value + " , hclase = " + numericUpDown4.Value + ", eqUtilizado = " + numericUpDown2.Value + " WHERE NumClase =  " + numericUpDown3.Value + " and idMateria = " + comboBox2.SelectedValue + "";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlCommand cmd = new SqlCommand(query, cn);
                cn.Open();
                cmd.ExecuteNonQuery();

            }

            LimpiarInterfazma();

        }

        //Boton para cerrar la ventana
        private void button5_Click(object sender, EventArgs e)
        {
            this.Close(); //Se utiliza la función .Close()
        }

        //Boton para eliminar datos de una tabla
        private void button10_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == -1) //Condicion que no permite eliminar datos a menos que se haya seleccionado una asignatura
            {
                MessageBox.Show("Seleccione una materia antes de eliminar datos");
                return;
            }

            //Se utiliza un DELETE FROM para eliminar los datos con la condición de seleccionar una materia del combobox3.
            string query = "DELETE FROM civregis WHERE idMateria = "+comboBox3.SelectedValue+" ";

            //Se manda a llamar la base de datos para eliminarlo directamente
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlDataAdapter da = new SqlDataAdapter(query, cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                cn.Open();
            }

        }




        //Boton para eliminar todos los datos de las tablas
        private void button11_Click(object sender, EventArgs e)
        {
                //Se utiliza un DELETE sin condicion para eliminar los datos
            string query = "DELETE FROM civregis";
            //Se manda a llamar la base de datos para eliminar los datos
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlDataAdapter da = new SqlDataAdapter(query, cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                cn.Open();
            }

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }


        //Boton que funciona para ver las horas totales del laboratorio
        private void button1_Click_1(object sender, EventArgs e)
        {
            panel4.Visible = true;
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            
            //Se hace una consulta desde la base de datos para contar el número de asignaturas
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("SELECT COUNT(asignatura) FROM civil ", cn);
                cn.Open();
                cmd.ExecuteNonQuery();

                int count = (int)cmd.ExecuteScalar();

                label20.Text = count.ToString();
                HorasTotales.Total6 = count;
                HorasTotales.Total18 = count;

            }
            
            //Se hace una consulta desde la base de datos para calcular el total de alumnos
            using (SqlConnection cn2 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("SELECT SUM(MaxAlumnos) FROM (SELECT MAX(NumAlumnos) AS MaxAlumnos FROM civregis GROUP BY idMateria) AS Subconsulta", cn2);
                cn2.Open();
                cmd.ExecuteNonQuery();

                int count = (int)cmd.ExecuteScalar();

                label19.Text = count.ToString();
                HorasTotales.Total5 = count;
                HorasTotales.Total17 = count;

            }

            //Se hace una consulta desde la base de datos para calcular el total de horas totales 
            using (SqlConnection cn3 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("SELECT sum(hclase) from civregis", cn3);
                cn3.Open();
                cmd.ExecuteNonQuery();

                decimal sum = Convert.ToDecimal(cmd.ExecuteScalar()); //Se convierten el resultado en decimal

                label18.Text = sum.ToString(); //Se agrega al label y se convierte en un String
                HorasTotales.Total4 = sum;
                HorasTotales.Total16 = sum;//Se acumulan las horas en la funcion horastotales en la variable correspondiente

            }





        }





        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }


        //Botones para calcular las horas dependiendo de asignaturas programadas, especiales e intersemestrales
        private void button12_Click(object sender, EventArgs e)
        {
            //Se manda a llamar a la base de datos
            using (SqlConnection cn4 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                try
                {
                    cn4.Open();
                    //Se agrega una consulta para calcular el total de horas dependiendo del rango de asignaturas
                    string query = "SELECT SUM(hclase) FROM civregis INNER JOIN civil ON civregis.idMateria = civil.idMateria WHERE civil.idMateria BETWEEN @materiaInicio AND @materiaFin";

                    using (SqlCommand cmd = new SqlCommand(query, cn4))
                    {
                        
                        int materiaInicio = Convert.ToInt32(comboBox7.SelectedValue); //Se convierten en enteros los rangos de la asignatura
                        int materiaFin = Convert.ToInt32(comboBox8.SelectedValue);


                        cmd.Parameters.AddWithValue("@materiaInicio", materiaInicio);
                        cmd.Parameters.AddWithValue("@materiaFin", materiaFin);


                        object result = cmd.ExecuteScalar();

                        //Una condicion que muestra que si el resultado es nulo entonces se mostrara en "0"
                        if (result != DBNull.Value && result != null)
                        {
                            decimal sum = Convert.ToDecimal(result); //El resultado se convierte en decimal
                            label21.Text = sum.ToString(); //El resultado se agrega al label y se convierte en un string
                            HorasTotales.Total = sum;
                            HorasTotales.Total13 = sum; //Se acumulan las horas en la función horasTotales en la variable correspondiente
                        }
                        else
                        {

                            label21.Text = "0";
                        }
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Error: " + ex.Message); //Si hay algun error se muestra en la terminal
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
           
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
                            HorasTotales.Total3 = sum;
                            HorasTotales.Total15 = sum;
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

        private void button15_Click(object sender, EventArgs e)
        {

        }

        private void button15_Click_1(object sender, EventArgs e)
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
                            HorasTotales.Total2 = sum;
                            HorasTotales.Total14 = sum;
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

        private DataTable cargardatos2()
        {
            DataTable dt = new DataTable();
            string query = "SELECT NumClase AS 'Número de clase', hentrada AS 'hora de entrada',hsalida AS 'hora de salida',NumAlumnos AS 'Numero de Alumnos',hclase AS 'Horas de clase',eqUtilizado AS 'Equipos Utilizados' FROM civil,civregis WHERE civil.idMateria = " + comboBox3.SelectedValue + " and civregis.idMateria = " + comboBox3.SelectedValue + "";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlDataAdapter da = new SqlDataAdapter(query, cn);
                da.Fill(dt);
            }
            return dt;
        }


        private void cargardatos()
        {
            
            string query = "SELECT NumClase AS 'Número de clase', hentrada AS 'hora de entrada',hsalida AS 'hora de salida',NumAlumnos AS 'Numero de Alumnos',hclase AS 'Horas de clase',eqUtilizado AS 'Equipos Utilizados' FROM civil,civregis WHERE civil.idMateria = " + comboBox3.SelectedValue + " and civregis.idMateria = " + comboBox3.SelectedValue + "";
            using (SqlConnection cn = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {

                SqlDataAdapter da = new SqlDataAdapter(query, cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                cn.Open();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument();

            int IC = 1;
            foreach(DataGridViewColumn column in dataGridView1.Columns)
            {
                sl.SetCellValue(1,IC, column.HeaderText.ToString());
                IC++;
                sl.AutoFitColumn(1,IC);
 
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
