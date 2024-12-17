using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Autolavado
{
    public partial class Cliente : Form
    {

        public Cliente()
        {
            InitializeComponent();
        }

        private readonly string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";
            
        private void Cliente_Load(object sender, EventArgs e)
        {
            LoadExcelData(excelFilePath);

            //Listar

            panel8.Hide(); //Modificar
            panel3.Hide(); //Agregar
            panel13.Hide(); //Eliminar
            panel1.Show(); //Listado
            panel16.Hide(); 
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Validar que no haya casillas vacias

            if (new[] { textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text }.Any(string.IsNullOrWhiteSpace))
            {
                MessageBox.Show("Ingrese informacion Valida");
                return;
            }

            // Validar que el mail contenga un correo electrónico válido
            string email = textBox5.Text;
            string emailPattern = @"^[^@\s]+@[^@\s]+\.[^@\s]+$";

            if (!Regex.IsMatch(email, emailPattern))
            {
                MessageBox.Show("Por favor, ingrese un correo electrónico válido.");
                return;
            }

            bool ciext = false;
            bool mailext = false;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                for (int row = 4; row <= 100; row++) // Lee hasta la fila 100
                {
                    string cellci = worksheet.Cells[row, 5].Text;
                    string inputci = textBox4.Text;

                    string cellmail = worksheet.Cells[row, 3].Text;
                    string inputmail = textBox5.Text;


                    if (cellci == inputci)
                    {
                        ciext = true;
                        MessageBox.Show("Usuario ya creado");
                    }

                    if (cellmail == inputmail)
                    {
                        mailext = true;
                        MessageBox.Show("Correo ya usado");
                    }
                }

                // Si ambos no existen, agregar una nueva fila
                if (!ciext && !mailext)
                {
                    string fr = membresia.nurandom();

                    int newRow = worksheet.Dimension.End.Row + 1;
                    worksheet.Cells[newRow, 1].Value = textBox2.Text; // Nombre
                    worksheet.Cells[newRow, 2].Value = textBox3.Text; // Apellido
                    worksheet.Cells[newRow, 3].Value = textBox5.Text; // Mail
                    worksheet.Cells[newRow, 4].Value = fr; // Membresia
                    worksheet.Cells[newRow, 5].Value = textBox4.Text; // CI

                    MessageBox.Show("Nuevo usuario agregado.");

                    MsgUtil.EnviarMembresia(textBox5.Text, fr);
                }
                

                package.Save();
            }            
        }



        private void button6_Click(object sender, EventArgs e)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                bool found = false;

                for (int row = 1; row <= 100; row++) // Lee hasta la fila 100
                {
                    string cellMembresia = worksheet.Cells[row, 4].Text; // Membresía en la columna 4
                    string inputMembresia = textBox10.Text; // Valor de la membresía que buscas

                    MessageBox.Show($"Comparando Membresía: '{cellMembresia}' con '{inputMembresia}'");

                    if (cellMembresia == inputMembresia)
                    {
                        found = true;
                        MessageBox.Show("Membresía encontrada, eliminando fila.");

                        // Elimina la fila
                        worksheet.DeleteRow(row);
                                                
                        break; // Sale del bucle después de eliminar la fila
                    }
                }
                                
                package.Save();
            }

        }

        private void LoadExcelData(string filePath)
        {
            var dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {

                var worksheet = package.Workbook.Worksheets[0]; // Primer hoja
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;


                // Agregar columnas al DataTable desde la columna C (3) hasta la columna H (8)
                for (int col = 1; col <= 6; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text); // Usar la fila 4 para los encabezados
                }

                // Rellena la tabla comenzando desde la fila 5
                // Cambia rowCount por el rango específico si es necesario

                for (int row = 2; row <= rowCount; row++) // Comienza desde la fila 5
                {
                    var newRow = dataTable.NewRow();
                    for (int col = 1; col <= 6; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text; // Ajusta el índice para el DataTable
                    }
                    dataTable.Rows.Add(newRow);
                }
            }
            // Asigna el DataTable al DataGridView
            dataGridView1.DataSource = dataTable;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            // Validar que el campo de búsqueda no esté vacío
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                LoadExcelData(excelFilePath);
                return;
            }

            string searchTerm = textBox1.Text; // Lo que se buscará
            int searchColumn = 4;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajamos con la primera hoja
                DataTable dt = new DataTable();

                // Crear columnas en el DataTable (Ajusta las columnas según tus necesidades)
                dt.Columns.Add("Nombre", typeof(string));
                dt.Columns.Add("Apellido", typeof(string));
                dt.Columns.Add("Mail", typeof(string));
                dt.Columns.Add("Membresia", typeof(string));
                dt.Columns.Add("CI", typeof(int));

                // Variable para controlar si se encontraron resultados
                bool foundResults = false;

                // Iterar sobre las filas de Excel desde la fila 4 hasta la última
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Verificar si el valor en la columna de búsqueda coincide con el término de búsqueda
                    if (worksheet.Cells[row, searchColumn].Text.Contains(searchTerm))
                    {
                        // Si se encuentra, crear una nueva fila en el DataTable con los valores de Excel
                        DataRow newRow = dt.NewRow();
                        newRow["Nombre"] = worksheet.Cells[row, 1].Text;
                        newRow["Apellido"] = worksheet.Cells[row, 2].Text;
                        newRow["Mail"] = worksheet.Cells[row, 3].Text;
                        newRow["Membresia"] = worksheet.Cells[row, 4].Text;
                        newRow["CI"] = int.Parse(worksheet.Cells[row, 5].Text);

                        dt.Rows.Add(newRow); // Agregar la fila encontrada al DataTable
                        foundResults = true; // Marcar que se encontró un resultado
                    }
                }

                // Asignar el DataTable al DataGridView para mostrar los resultados
                dataGridView1.DataSource = dt;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Modificar

            panel8.Hide(); //Modificar
            panel3.Hide(); //Agregar
            panel13.Hide(); //Eliminar
            panel1.Hide(); //Listado
            panel16.Show(); //COdigoModificar
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Agregar

            panel8.Hide(); //Modificar
            panel3.Show(); //Agregar
            panel13.Hide(); //Eliminar
            panel1.Hide(); //Listado
            panel16.Hide(); //COdigoModificar

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Eliminar

            panel8.Hide(); //Modificar
            panel3.Hide(); //Agregar
            panel13.Show(); //Eliminar
            panel1.Hide(); //Listado
            panel16.Hide(); //COdigoModificar

        }

        private int foundRow = -1;
        private void button7_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Autolavado.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            foundRow = -1; // Reiniciar la fila encontrada

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajamos con la primera hoja

                // Iterar sobre las filas desde la fila 4 hasta la última
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Buscar en la columna 4 el código (ISBN)
                    if (worksheet.Cells[row, 4].Text == textBox11.Text)
                    {
                        foundRow = row; // Guardar la fila donde se encontró el código
                        MessageBox.Show("Código encontrado en la fila: " + row); // Mostrar la fila donde se encontró
                        panel8.Show();
                        panel16.Hide();

                        break; // Salir del bucle si se encontró el código
                                                
                    }
                }
            }

            if (foundRow == -1)
            {
                MessageBox.Show("El código no fue encontrado en la tabla de Excel.");
            }
        }

        private void ModifyRowInExcel(int row)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Autolavado.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajamos con la primera hoja

                // Si la fila es válida (es decir, el código fue encontrado), modificar la fila
                if (row > 0)
                {
                    // Solo se actualizan las celdas si los TextBox tienen información
                    if (!string.IsNullOrWhiteSpace(textBox9.Text)) worksheet.Cells[row, 1].Value = textBox9.Text; // Columna 1: Nombre
                    if (!string.IsNullOrWhiteSpace(textBox7.Text)) worksheet.Cells[row, 2].Value = textBox7.Text; // Columna 2: Apellido
                    if (!string.IsNullOrWhiteSpace(textBox8.Text)) worksheet.Cells[row, 3].Value = textBox8.Text; // Columna 3: Mail
                    if (!string.IsNullOrWhiteSpace(textBox6.Text)) worksheet.Cells[row, 5].Value = textBox6.Text; // Columna 5: CI
                    

                    // Guardar los cambios en el archivo Excel
                    FileInfo file = new FileInfo(excelFilePath);
                    package.SaveAs(file);

                    // Mostrar mensaje de confirmación
                    MessageBox.Show("La fila ha sido modificada correctamente.");
                }
                else
                {
                    MessageBox.Show("No se encontró la fila para modificar.");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (foundRow > 0)
            {
                ModifyRowInExcel(foundRow); // Modificar la fila con los datos de los TextBox
            }
            else
            {
                MessageBox.Show("Primero debes buscar el código antes de modificar.");
            }
        }
    }
}
