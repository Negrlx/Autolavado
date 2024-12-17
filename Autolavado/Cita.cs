using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace Autolavado
{
    public partial class Cita : Form
    {

        public Cita()
        {
            InitializeComponent();
        }

        private readonly string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

        private void Cita_Load(object sender, EventArgs e)
        {
            LoadExcelData(excelFilePath);
            CargarColaDesdeExcel();

            radioButton1.Checked = true;

            //Listar

            panel1.Show(); //Lista
            panel3.Hide(); //Agregar
            panel13.Hide(); //Eliminar

        }
        private void LoadExcelData(string filePath)
        {
            var dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {

                var worksheet = package.Workbook.Worksheets[1]; 
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
            dataGridView1.DataSource = dataTable;
        }

        //Creacion de Colas - Pila - Arreglo

        public Cola climpieza = new Cola();
        public Cola caceite = new Cola();
        public Cola cbalanceo = new Cola();

        public Cola climpiezaAux = new Cola();
        public Cola caceiteAux = new Cola();
        public Cola cbalanceoAux = new Cola();

        public Cola cgeneral = new Cola();
        public Cola cgeneralEx = new Cola();


        public Pila<string> cauchos = new Pila<string>();

        public ElementoCola carro;

        private void CargarColaDesdeExcel() // Cargar datos de Excel a las colas
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet sheet0 = package.Workbook.Worksheets[1]; // Hoja 0
                int totalRows = sheet0.Dimension.Rows; // Obtener el número total de filas

                for (int row = 2; row <= totalRows; row++) // Empezamos desde la fila 2 (omitiendo encabezado)
                {
                    // Obtener los datos de la fila
                    string vehiculo = sheet0.Cells[row, 1].Text;   // Columna 1: Vehiculo
                    string modelo = sheet0.Cells[row, 2].Text;     // Columna 2: Modelo
                    string placa = sheet0.Cells[row, 3].Text;      // Columna 3: Placa
                    string servicio = sheet0.Cells[row, 4].Text;   // Columna 4: Servicio
                    string membresia = sheet0.Cells[row, 6].Text;  // Columna 6: Membresia

                    // Crear un nuevo ElementoCola basado en los datos
                    ElementoCola elemento = new ElementoCola(vehiculo, modelo, placa, membresia, servicio);

                    // Asignar el servicio a la cola correspondiente
                    if (servicio == "Aseo-Aspirado-Secado")
                    {
                        climpieza.Insertar(elemento);
                        climpiezaAux.Insertar(elemento);
                    }
                    else if (servicio == "Cambio-Aceite")
                    {
                        caceite.Insertar(elemento);
                        caceiteAux.Insertar(elemento);
                    }
                    else if (servicio == "Balanceo")
                    {
                        // Si el servicio es Balanceo, también se asigna la pila de cauchos
                        int cant = 4; // Moto o Carro
                        Pila<int> cauchos = new Pila<int>();
                        for (int i = 0; i < cant; i++)
                        {
                            cauchos.Push(i);
                        }

                        elemento.AsignarPila(cauchos);

                        cbalanceo.Insertar(elemento);
                        cbalanceoAux.Insertar(elemento);
                    }
                }

                MessageBox.Show("Los datos han sido cargados correctamente a las colas.");
            }
        }
        private void button5_Click(object sender, EventArgs e) //Agregar Elemento a las Colas 
        {
            if (new[] { textBox3.Text, textBox4.Text, textBox5.Text }.Any(string.IsNullOrWhiteSpace))
            {
                MessageBox.Show("Todos los campos deben estar llenos");
                return;
            }

            if (!(radioButton1.Checked ^ radioButton2.Checked ^ radioButton3.Checked)) // XOR asegura que solo uno esté activado
            {
                MessageBox.Show("Debe seleccionar un servicio");
                return;
            }

            if (!(radioButton4.Checked ^ radioButton5.Checked))
            {
                MessageBox.Show("Debe seleccionar un tipo de Vehículo");
                return;
            }

            string placaIngresada = textBox4.Text;
            string membresiaIngresada = textBox5.Text;

            bool membresiaEncontrada = false;
            bool placaRepetida = false;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet sheet0 = package.Workbook.Worksheets[0]; // Hoja 0
                ExcelWorksheet sheet1 = package.Workbook.Worksheets[1]; // Hoja 1 (si corresponde)

                // Recorremos todas las filas de la hoja 0
                for (int row = 2; row <= sheet0.Dimension.End.Row; row++) // Comenzamos desde la fila 2 (asumiendo encabezados)
                {
                    string placaExistente = sheet1.Cells[row, 3].Text.Trim();    // Columna 3: Placa
                    string membresiaExistente = sheet0.Cells[row, 4].Text.Trim(); // Columna 6: Membresía

                    if (membresiaExistente == membresiaIngresada)
                    {
                        membresiaEncontrada = true;

                        if (placaExistente == placaIngresada)
                        {
                            placaRepetida = true;
                            break; // No necesitamos seguir buscando
                        }
                    }
                }

                
            }

            if (!membresiaEncontrada)
            {
                MessageBox.Show("La membresía ingresada no se encuentra en la base de datos. Regístrese.");
                return;
            }

            if (placaRepetida)
            {
                MessageBox.Show("La placa ingresada ya está asociada a esta membresía. Use una placa diferente.");
                return;
            }

            // Continuar con la lógica de inserción...
            string servicio = radioButton1.Checked ? "Aseo-Aspirado-Secado" :
                              radioButton2.Checked ? "Cambio-Aceite" :
                              "Balanceo";

            string vehiculo = radioButton4.Checked ? "Carro" : "Camioneta";

            carro = new ElementoCola(vehiculo, textBox3.Text, placaIngresada, membresiaIngresada, servicio);

            if (servicio == "Aseo-Aspirado-Secado")
            {
                climpieza.Insertar(carro);
                climpiezaAux.Insertar(carro);
                MessageBox.Show("Elemento agregado a la cola de limpieza.");
                MessageBox.Show($"Cantidad de la Cola Limpieza: {climpieza.Cantidad()}");
            }
            else if (servicio == "Cambio-Aceite")
            {
                caceite.Insertar(carro);
                caceiteAux.Insertar(carro);
                MessageBox.Show("Elemento agregado a la cola de aceite.");
                MessageBox.Show($"Cantidad de la Cola Aceite: {caceite.Cantidad()}");
            }
            else if (servicio == "Balanceo")
            {
                int cant = 4; // Moto o Carro
                Pila<int> cauchos = new Pila<int>();
                for (int i = 0; i < cant; i++)
                {
                    cauchos.Push(i);
                }

                carro.AsignarPila(cauchos);

                cbalanceo.Insertar(carro);
                cbalanceoAux.Insertar(carro);
                MessageBox.Show("Elemento agregado a la cola de balanceo.");
                MessageBox.Show($"Cantidad de la Cola Balanceo: {cbalanceo.Cantidad()}");
            }
        }
        private void button6_Click(object sender, EventArgs e) //Eliminar Elemento de las Colas
        {
            string membresia = textBox10.Text;
            bool encontrado = false;

            // Intentar eliminar de la cola de limpieza
            if (climpieza.BuscarPosicionMembresia(membresia) != -1)
            {
                climpieza.EliminarElementoPorMembresia(membresia);
                encontrado = true;
            }
            // Intentar eliminar de la cola de aceite
            else if (caceite.BuscarPosicionMembresia(membresia) != -1)
            {
                caceite.EliminarElementoPorMembresia(membresia);
                encontrado = true;
            }
            // Intentar eliminar de la cola de balanceo
            else if (cbalanceo.BuscarPosicionMembresia(membresia) != -1)
            {
                cbalanceo.EliminarElementoPorMembresia(membresia);
                encontrado = true;
            }

            // Mostrar el resultado
            if (!encontrado)
            {
                MessageBox.Show("No se eliminó la cita a nombre de esa membresía.");
            }
            else
            {
                MessageBox.Show("Se encontró la cita a nombre de esa membresía.");
                MessageBox.Show("Cita eliminada.");
            }
        }

        private void button2_Click(object sender, EventArgs e) //Plantilla Agregar
        {   

            panel1.Hide(); //Lista
            panel3.Show(); //Agregar
            panel13.Hide(); //Eliminar
        }
        private void button3_Click(object sender, EventArgs e) //Plantilla Eliminar
        {
            
            panel1.Hide(); //Lista
            panel3.Hide(); //Agregar
            panel13.Show(); //Eliminar
        }
        public void MezclarColas()
        {
            bool tbr = false;
            int cant = 0;

            // Bucle que continúa mientras al menos una cola auxiliar tenga elementos
            while (!climpiezaAux.EsVacia() || !caceiteAux.EsVacia() || !cbalanceoAux.EsVacia())
            {
                ElementoCola elemento;

                // Retirar elementos de cada cola auxiliar y agregarlos a las colas generales
                if (!climpiezaAux.EsVacia())
                {
                    elemento = climpiezaAux.Retirar();
                    cgeneralEx.Insertar(elemento);
                    cgeneral.Insertar(elemento);
                    cant++;
                }

                if (!caceiteAux.EsVacia())
                {
                    elemento = caceiteAux.Retirar();
                    cgeneralEx.Insertar(elemento);
                    cgeneral.Insertar(elemento);
                    cant++;
                }

                if (!cbalanceoAux.EsVacia())
                {
                    elemento = cbalanceoAux.Retirar();
                    cgeneralEx.Insertar(elemento);
                    cgeneral.Insertar(elemento);
                    cant++;
                }

                tbr = true;
            }

            // Mostrar mensaje después de mezclar
            if (tbr)
            {
                MessageBox.Show("Se agregaron todas las colas particulares para generalizar");
                MessageBox.Show($"Se agregaron: {cant} elementos a la Cola");
            }
            else
            {
                MessageBox.Show("No se agregaron colas particulares para generalizar");
            }
        }

        private void EliminarTodasLasFilas()
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[1]; // Hoja 0 (puedes cambiar el índice si es necesario)

                int totalRows = sheet.Dimension.Rows;

                if (totalRows > 1)
                {
                    sheet.DeleteRow(2, totalRows - 1); // Elimina las filas desde la 2 hasta la última
                }
                package.Save();

                MessageBox.Show("Todas las filas han sido eliminadas correctamente.");
            }
        }
        private void CargarColaAExcel() //Subir la cola GeneralEx al Excel
        {
            EliminarTodasLasFilas();    
            if (!cgeneralEx.EsVacia())
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet sheet1 = package.Workbook.Worksheets[1]; 

                    int newRow = sheet1.Dimension?.End.Row + 1 ?? 1;

                    do
                    {
                        ElementoCola elemento = cgeneralEx.Retirar(); // Retiramos el primer elemento de la cola

                        // Agregar los datos del ElementoCola a las celdas correspondientes de la hoja 1
                        sheet1.Cells[newRow, 1].Value = elemento.Vehiculo; // Vehiculo
                        sheet1.Cells[newRow, 2].Value = elemento.Modelo;   // Modelo
                        sheet1.Cells[newRow, 3].Value = elemento.Placa;    // Placa
                        sheet1.Cells[newRow, 6].Value = elemento.Membresia; // Membresia
                        sheet1.Cells[newRow, 4].Value = elemento.Servicio; // Servicio


                        // Verificamos si tiene pila asociada y agregamos la cantidad
                        int pilaCantidad = elemento.PilaOpcional != null ? elemento.PilaOpcional.Count() : 0;
                        sheet1.Cells[newRow, 5].Value = pilaCantidad;      // Cantidad de la pila

                        newRow++; // Incrementar la fila para el siguiente elemento

                    } while (!cgeneralEx.EsVacia());
                    // Guardar los cambios
                    package.Save();
                }

                MessageBox.Show("Los datos de la cola se han agregado correctamente a la hoja 1.");
            }
            else
            {
                MessageBox.Show("Los datos de la cola NO se han agregado correctamente a la hoja 1.");

            }

        }
        private void Cita_FormClosing(object sender, FormClosingEventArgs e)
        {
            MezclarColas();
            CargarColaAExcel();
        }
    }
}


//Cargar Datos del excel a las Colas y al Grid
//Manejar las Colas
//Borrar Excel
//Subir las Colas al Excel