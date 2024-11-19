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

namespace Autolavado
{
    public partial class Consulta : Form
    {

        public Cola climpieza = new Cola(10);
        public Cola caceite = new Cola(5);
        public Cola cbalanceo = new Cola(5);

        public Cola cgeneral = new Cola(20);

        public Pila cauchos = new Pila();

        public ElementoCola carro;

        private readonly string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

        private void LoadExcelData(string filePath)
        {
            // Crear DataTables para cada Grid
            var dataTable1 = new DataTable(); // Para Grid1
            var dataTable3 = new DataTable(); // Para Grid3
            var dataTable4 = new DataTable(); // Para Grid4

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[1];
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                // Agregar columnas a los DataTables
                // Para Grid1 (Servicio = "Aseo-Aspirado-Secado")
                dataTable1.Columns.Add("Vehiculo");
                dataTable1.Columns.Add("Modelo");
                dataTable1.Columns.Add("Placa");
                dataTable1.Columns.Add("Membresia");

                // Para Grid3 (Servicio = "Cambio-Aceite")
                dataTable3.Columns.Add("Vehiculo");
                dataTable3.Columns.Add("Modelo");
                dataTable3.Columns.Add("Placa");
                dataTable3.Columns.Add("Membresia");

                // Para Grid4 (Servicio = "Balanceo")
                dataTable4.Columns.Add("Vehiculo");
                dataTable4.Columns.Add("Modelo");
                dataTable4.Columns.Add("Placa");
                dataTable4.Columns.Add("Cauchos At");
                dataTable4.Columns.Add("Membresia");

                // Iterar sobre las filas del Excel
                for (int row = 2; row <= rowCount; row++) // Comenzamos desde la fila 2 para evitar el encabezado
                {
                    string servicio = worksheet.Cells[row, 4].Text; // Columna 4: "Servicio"

                    if (servicio == "Aseo-Aspirado-Secado")
                    {
                        // Agregar solo las columnas 1, 2, 3 y 6 al DataTable de Grid1
                        var newRow = dataTable1.NewRow();
                        newRow[0] = worksheet.Cells[row, 1].Text; // Vehiculo
                        newRow[1] = worksheet.Cells[row, 2].Text; // Modelo
                        newRow[2] = worksheet.Cells[row, 3].Text; // Placa
                        newRow[3] = worksheet.Cells[row, 6].Text; // Membresia
                        dataTable1.Rows.Add(newRow);
                    }
                    else if (servicio == "Balanceo")
                    {
                        // Agregar solo las columnas 1, 2, 3, 5 y 6 al DataTable de Grid4
                        var newRow = dataTable4.NewRow();
                        newRow[0] = worksheet.Cells[row, 1].Text; // Vehiculo
                        newRow[1] = worksheet.Cells[row, 2].Text; // Modelo
                        newRow[2] = worksheet.Cells[row, 3].Text; // Placa
                        newRow[3] = worksheet.Cells[row, 5].Text; // Membresia
                        newRow[4] = worksheet.Cells[row, 6].Text; // Membresia
                        dataTable4.Rows.Add(newRow);
                    }
                    else if (servicio == "Cambio-Aceite")
                    {
                        // Agregar solo las columnas 1, 2, 3 y 6 al DataTable de Grid3
                        var newRow = dataTable3.NewRow();
                        newRow[0] = worksheet.Cells[row, 1].Text; // Vehiculo
                        newRow[1] = worksheet.Cells[row, 2].Text; // Modelo
                        newRow[2] = worksheet.Cells[row, 3].Text; // Placa
                        newRow[3] = worksheet.Cells[row, 6].Text; // Membresia
                        dataTable3.Rows.Add(newRow);
                    }
                }
            }

            // Asignar los DataTables a los respectivos DataGridViews
            dataGridView1.DataSource = dataTable1;  // Para el servicio "Aseo-Aspirado-Secado"
            dataGridView3.DataSource = dataTable3;  // Para el servicio "Cambio-Aceite"
            dataGridView4.DataSource = dataTable4;  // Para el servicio "Balanceo"
        }

        public Consulta()
        {
            InitializeComponent();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void Consulta_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = true;

            LoadExcelData(excelFilePath);
            CargarListos(excelFilePath);
        }

        private void CargarListos(string filePath)
        {
            var dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {

                var worksheet = package.Workbook.Worksheets[2];
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;


                for (int col = 1; col <= 6; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var newRow = dataTable.NewRow();
                    for (int col = 1; col <= 6; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(newRow);
                }
            }
            dataGridView2.DataSource = dataTable;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //Mostrar Limpieza-Aspirado-Secada 

            dataGridView1.Show();
            dataGridView3.Hide();
            dataGridView4.Hide();
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //Mostrar Cambio-Aceite

            dataGridView1.Hide();
            dataGridView3.Show();
            dataGridView4.Hide();
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            //Mostrar Balanceo

            dataGridView1.Hide();
            dataGridView3.Hide();
            dataGridView4.Show();
        }

        private void button4_Click(object sender, EventArgs e) //Lista de Espera
        {
            panel1.Hide(); //Inicio
            panel13.Hide(); //Personalizada
            panel2.Show(); //Espera
            panel3.Hide(); //Listos
        }
        private void button2_Click(object sender, EventArgs e) //Especializada
        {
            panel1.Hide(); //Inicio
            panel13.Show(); //Personalizada
            panel2.Hide(); //Espera
            panel3.Hide(); //Listos
        }
        private void button1_Click(object sender, EventArgs e) //Listos
        {
            panel1.Hide(); //Inicio
            panel13.Hide(); //Personalizada
            panel2.Hide(); //Espera
            panel3.Show(); //Listos
        }
        private void button6_Click(object sender, EventArgs e) //Busqueda Personalizada
        {
            bool found = false; // Variable para verificar si encontramos la coincidencia

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                // Buscar en la Hoja 1
                ExcelWorksheet sheet0 = package.Workbook.Worksheets[1]; // Hoja 1
                int lastRowSheet0 = sheet0.Dimension.End.Row; // Obtener la última fila con datos de la Hoja 1

                for (int row = 2; row <= lastRowSheet0; row++) // Empezamos desde la fila 2 (omitiendo encabezado)
                {
                    string vehiculo = sheet0.Cells[row, 1].Text;
                    string modelo = sheet0.Cells[row, 2].Text;
                    string placa = sheet0.Cells[row, 3].Text;
                    string servicio = sheet0.Cells[row, 4].Text;
                    string membresia = sheet0.Cells[row, 6].Text;
                    string estado = "En Lista de Espera";



                    if (membresia == textBox10.Text)
                    {                        
                        MessageBox.Show($"Vehiculo del Cliente: {vehiculo}.\nModelo del Vehiculo: {modelo}.\nPlaca del Vehiculo: {placa}.\nServicio para el Vehiculo: {servicio}.\nMembresia del Cliente: {membresia}.\nEstado del Vehiculo: {estado}.");

                        found = true;
                        break; // Salir del bucle una vez encontrada la coincidencia
                    }
                }

                // Si no se encontró en la Hoja 1, buscar en la Hoja 2
                if (!found)
                {
                    ExcelWorksheet sheet1 = package.Workbook.Worksheets[2]; // Hoja 2
                    int lastRowSheet1 = sheet1.Dimension.End.Row; // Obtener la última fila con datos de la Hoja 2

                    for (int row = 2; row <= lastRowSheet1; row++) // Empezamos desde la fila 2 (omitiendo encabezado)
                    {
                        string vehiculo = sheet0.Cells[row, 1].Text;
                        string modelo = sheet0.Cells[row, 2].Text;
                        string placa = sheet0.Cells[row, 3].Text;
                        string servicio = sheet0.Cells[row, 4].Text;
                        string membresia = sheet0.Cells[row, 6].Text;
                        string estado = "Ya procesado";

                        if (membresia.Equals(textBox10.Text, StringComparison.OrdinalIgnoreCase))
                        {
                            MessageBox.Show($"Vehiculo del Cliente: {vehiculo}.\nModelo del Vehiculo: {modelo}.\nPlaca del Vehiculo: {placa}.\nServicio para el Vehiculo: {servicio}.\nMembresia del Cliente: {membresia}.\nEstado del Vehiculo: {estado}.");
                            
                            found = true;
                            break; // Salir del bucle una vez encontrada la coincidencia
                        }
                    }
                }

                // Si no se encuentra en ninguna hoja
                if (!found)
                {
                    MessageBox.Show("No se encontró una coincidencia para la membresía en ninguna de las hojas.");
                }
            }
        }


    }
}
