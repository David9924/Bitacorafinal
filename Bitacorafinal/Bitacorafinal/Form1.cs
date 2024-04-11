using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using LinqForEEPlus;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Bitacorafinal
{
    public partial class Form1 : Form
    {
        private bool form2Abierto = false;
        private bool form3Abierto = false;

        public Form1()
        {
            InitializeComponent();
            panel4.Visible = false;
            panel1.Visible = false;
            panel3.Visible = false;

        }

        //Boton para abrir las estadisticas de especialidades de civiles
        private void button1_Click(object sender, EventArgs e)
        {
            if (!form2Abierto)
            {
                Form2 form2 = new Form2();
                form2.Show(); //Se muestra el form 2 
                form2Abierto = true;

                form2.FormClosed += (s, args) => form2Abierto = false; //Abre una sola ventana. 

            }
        }

        //Boton para abrir las estadisticas de Geomatica
        private void button2_Click(object sender, EventArgs e)
        {
            if (!form3Abierto)
            {
                Form3 form3 = new Form3();
                form3.Show(); //Se muestra el form 3
                form3Abierto = true;

                form3.FormClosed += (s, args) => form3Abierto = false; //Abre una sola ventana

            }

        }

        //Boton para cerrar la ventana
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Boton para mostrar las horas totales
        private void button4_Click(object sender, EventArgs e)
        {
            panel4.Visible = true; //Se visualiza uno de los paneles
            //Se ocultan los botones
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button13.Visible = false;

        }

        //Boton para regresar
        private void button5_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            button1.Visible = true;
            button2.Visible = true;
            button3.Visible = true;
            button4.Visible = true;
            button13.Visible = true;

        }

        //Boton para actualizar los datos
        private void button6_Click(object sender, EventArgs e)
        {
            Horas();
        }

        //Se descargan dato de horas totales
        private void button7_Click(object sender, EventArgs e)
        {
            GenerarArchivo();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel1.Visible = true;

        }

        private void button11_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
            panel1.Visible = false;

        }

        private void button15_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel1.Visible = true;

        }

        private void button9_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel1.Visible = false;
            panel3.Visible = true;
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel1.Visible = true;
            panel3.Visible = false;
        }

        //Guardar
        private void button10_Click(object sender, EventArgs e)
        {
            GenerarArchivo();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            GenerarArchivo();
        }

        //Funcion para generar documento 
        private void GenerarArchivo()
        {
            // Crear un nuevo diálogo de guardar archivo
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Guardar datos de la materia";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)

            {
                string filePath = saveFileDialog.FileName;
                string[] etiquetas = { "Horas totales de asignaturas programadas", "Horas totales de asignaturas adicionales", "Horas totales de cursos intersemestrales", "Horas totales de uso", "Numero de Alumnos", "Numero de asignaturas" };
                string[] valores = { label21.Text, label2.Text, label3.Text, label4.Text, label5.Text, label6.Text };
                string[] etiquetas2 = { "Horas totales de asignaturas programadas de Geomatica", "Horas totales de asignaturas adicionales de Geomatica", "Horas totales de cursos intersemestrales de Geomatica", "Horas totales de uso de Geomatica", "Numero de Alumnos de Geomatica", "Numero de asignaturas de Geomatica" };
                string[] valores2 = { label19.Text, label18.Text, label11.Text, label10.Text, label9.Text, label8.Text };
                string[] etiquetas3 = { "Horas totales de asignaturas programadas de Civil", "Horas totales de asignaturas adicionales de Civil", "Horas totales de cursos intersemestrales de Civil", "Horas totales de uso de Civil", "Numero de Alumnos de Civil", "Numero de asignaturas de Civil" };
                string[] valores3 = { label33.Text, label32.Text, label31.Text, label30.Text, label29.Text, label28.Text };


                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Si el archivo Excel ya existe, eliminar la hoja "Datos"
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Datos");
                    if (worksheet != null)
                    {
                        package.Workbook.Worksheets.Delete(worksheet);
                    }

                    // Crear una nueva hoja "Datos"
                    worksheet = package.Workbook.Worksheets.Add("Datos");

                    //Se da formato a la celda
                    var formato = worksheet.Cells["A1"].Style; //Se selecciona el rango de celdas para darle estilo
                    formato.Font.Bold = true; //Se agrega negrita
                    formato.Font.Size = 12; //Tamaño de fuente
                                            //Se agrega color a la celda
                    formato.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; //Relleno solido
                    formato.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); //Color de la celda
                                                                                           //Se agregan bordes
                    formato.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    //Se centra el texto
                    formato.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    formato.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;



                    //Se da formato a la celdas

                    var formato2 = worksheet.Cells["A2:B7"].Style; //Se selecciona el rango de celdas para darle estilo
                    formato2.Font.Size = 12; //Tamaño de fuente

                    //Se agregan bordes interiores y exteriores
                    formato2.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    formato2.Border.Left.Style = ExcelBorderStyle.Thin;
                    formato2.Border.Right.Style = ExcelBorderStyle.Thin;
                    formato2.Border.Top.Style = ExcelBorderStyle.Thin;
                    formato2.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Se agrega el texto a la celda A1
                    worksheet.Cells["A1"].Value = "Horas totales de Laboratorio de Geomatica y Especialidades de Civiles";


                    //Se da formato a la celda
                    var formato3 = worksheet.Cells["A9"].Style; //Se selecciona el rango de celdas para darle estilo
                    formato3.Font.Bold = true; //Se agrega negrita
                    formato3.Font.Size = 12; //Tamaño de fuente
                                             //Se agrega color a la celda
                    formato3.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; //Relleno solido
                    formato3.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); //Color de la celda
                                                                                            //Se agregan bordes
                    formato3.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    //Se centra el texto
                    formato3.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    formato3.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["A9"].Value = "Horas totales de Laboratorio de Geomatica";

                    //Se da formato a la celdas

                    var formato4 = worksheet.Cells["A10:B15"].Style; //Se selecciona el rango de celdas para darle estilo
                    formato4.Font.Size = 12; //Tamaño de fuente

                    //Se agregan bordes interiores y exteriores
                    formato4.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    formato4.Border.Left.Style = ExcelBorderStyle.Thin;
                    formato4.Border.Right.Style = ExcelBorderStyle.Thin;
                    formato4.Border.Top.Style = ExcelBorderStyle.Thin;
                    formato4.Border.Bottom.Style = ExcelBorderStyle.Thin;


                    //Formato para civil

                    //Se da formato a la celda
                    var formato5 = worksheet.Cells["A17"].Style; //Se selecciona el rango de celdas para darle estilo
                    formato5.Font.Bold = true; //Se agrega negrita
                    formato5.Font.Size = 12; //Tamaño de fuente
                                             //Se agrega color a la celda
                    formato5.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; //Relleno solido
                    formato5.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); //Color de la celda
                                                                                            //Se agregan bordes
                    formato5.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    //Se centra el texto
                    formato5.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    formato5.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["A17"].Value = "Horas totales de Laboratorio de Especialidades de Civiles";

                    var formato6 = worksheet.Cells["A18:B23"].Style; //Se selecciona el rango de celdas para darle estilo
                    formato4.Font.Size = 12; //Tamaño de fuente

                    //Se agregan bordes interiores y exteriores
                    formato6.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    formato6.Border.Left.Style = ExcelBorderStyle.Thin;
                    formato6.Border.Right.Style = ExcelBorderStyle.Thin;
                    formato6.Border.Top.Style = ExcelBorderStyle.Thin;
                    formato6.Border.Bottom.Style = ExcelBorderStyle.Thin;


                    // Escribir los encabezados
                    for (int i = 0; i < etiquetas.Length; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = etiquetas[i];
                    }
                    // Escribir los valores
                    for (int i = 0; i < valores.Length; i++)
                    {
                        worksheet.Cells[i + 2, 2].Value = valores[i];
                    }

                    //Etiqueta de Geomatica

                    for (int i = 0; i < etiquetas2.Length; i++)
                    {
                        worksheet.Cells[i + 10, 1].Value = etiquetas2[i];
                    }

                    // Escribir los valores
                    for (int i = 0; i < valores2.Length; i++)
                    {
                        worksheet.Cells[i + 10, 2].Value = valores2[i];
                    }

                    //Etiquetas de civil

                    for (int i = 0; i < etiquetas3.Length; i++)
                    {
                        worksheet.Cells[i + 18, 1].Value = etiquetas3[i];
                    }

                    // Escribir los valores
                    for (int i = 0; i < valores3.Length; i++)
                    {
                        worksheet.Cells[i + 18, 2].Value = valores3[i];
                    }

                    worksheet.Cells.AutoFitColumns();

                    // Guardar el archivo Excel
                    package.Save();
                }



                MessageBox.Show("Datos guardados exitosamente", "Datos Guardados", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }


        }

        private void Horas()
        {
            //Horas totales de Especialidades de civiles y Geomatica
            label21.Text = HorasTotales.Total.ToString(); //Se manda a llamar la clase HorasTotales acumuladas en el label
            label2.Text = HorasTotales.Total2.ToString();
            label3.Text = HorasTotales.Total3.ToString();
            label4.Text = HorasTotales.Total4.ToString();
            label5.Text = HorasTotales.Total5.ToString();
            label6.Text = HorasTotales.Total6.ToString();

            //Horas totales de Geomatica
            label19.Text = HorasTotales.Total7.ToString();
            label18.Text = HorasTotales.Total8.ToString();
            label11.Text = HorasTotales.Total9.ToString();
            label10.Text = HorasTotales.Total10.ToString();
            label9.Text = HorasTotales.Total11.ToString();
            label8.Text = HorasTotales.Total12.ToString();

            //Horas totales de civiles
            label33.Text = HorasTotales.Total13.ToString();
            label32.Text = HorasTotales.Total14.ToString();
            label31.Text = HorasTotales.Total15.ToString();
            label30.Text = HorasTotales.Total16.ToString();
            label29.Text = HorasTotales.Total17.ToString();
            label28.Text = HorasTotales.Total18.ToString();

        }

        private void button12_Click(object sender, EventArgs e)
        {
            Horas();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Horas();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            DialogResult advertencia = MessageBox.Show("Se borrara el contenido de las bitacoras ¿Estas seguro?", "Eliminación de datos", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (advertencia == DialogResult.Yes)
            {
                borrardatos();
                MessageBox.Show("Los datos han sido borrados correctamente.", "Borrado Exitoso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El borrado ha sido cancelado.", "Borrado Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void borrardatos()
        {
            using (SqlConnection cn1 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("Delete from civregis", cn1);
                cn1.Open();
                cmd.ExecuteNonQuery();

            }
            using (SqlConnection cn2 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("Delete from civil", cn2);
                cn2.Open();
                cmd.ExecuteNonQuery();

            }

            using (SqlConnection cn3 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("Delete from georegis", cn3);
                cn3.Open();
                cmd.ExecuteNonQuery();

            }

            using (SqlConnection cn4 = new SqlConnection("Data Source=DESKTOP-UV70V09\\SQLEXPRESS;Initial Catalog=bitacora;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand("Delete from geo", cn4);
                cn4.Open();
                cmd.ExecuteNonQuery();

            }


        }
    }


    //Clase que permite calcular las horas totales de los dos laboratorios
    public static class HorasTotales
    {
        //Contador de horas totales
        public static decimal Total { get; set; }
        public static decimal Total2 { get; set; }
        public static decimal Total3 { get; set; }
        public static decimal Total4 { get; set; }
        public static decimal Total5 { get; set; }
        public static decimal Total6 { get; set; }

        //Contador de horas totales geomatica
        public static decimal Total7 { get; set; }
        public static decimal Total8 { get; set; }
        public static decimal Total9 { get; set; }
        public static decimal Total10 { get; set; }
        public static decimal Total11 { get; set; }
        public static decimal Total12 { get; set; }

        //Contador de horas totales de geomatica
        public static decimal Total13 { get; set; }
        public static decimal Total14 { get; set; }
        public static decimal Total15 { get; set; }
        public static decimal Total16 { get; set; }
        public static decimal Total17 { get; set; }
        public static decimal Total18 { get; set; }

    }


    //Clase que permite escribir los datos en un archivo de excel.
    public class Excel
    {
        public static void Escribir(string filePath, string[] etiquetas, string[] valores)
        {
            FileInfo newFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Datos");

                for(int i = 0; i < etiquetas.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = etiquetas[i];
                }

                for (int i = 0; i < valores.Length; i++)
                {
                    worksheet.Cells[2, i + 1].Value = valores[i];
                }

                package.Save();
            }
        }
    }


}
