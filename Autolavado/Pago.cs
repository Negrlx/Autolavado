using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Autolavado
{
    public partial class Pago : Form
    {
        public Pago()
        {
            InitializeComponent();
        }

        private ElementoCola bm(string inputMembresia)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                for (int row = 1; row <= 100; row++) // Lee hasta la fila 100
                {
                    string cellMembresia = worksheet.Cells[row, 6].Text; // Membresía en la columna 4

                    Debug.WriteLine($"Comparando Membresía: '{cellMembresia}' con '{inputMembresia}'");

                    if (cellMembresia == inputMembresia)
                    {
                        Debug.WriteLine("Membresía encontrada, obteniendo valor de la columna 3.");

                        string vehiculo = worksheet.Cells[row, 1].Text;   // Columna 1: Vehiculo
                        string modelo = worksheet.Cells[row, 2].Text;     // Columna 2: Modelo
                        string placa = worksheet.Cells[row, 3].Text;      // Columna 3: Placa
                        string servicio = worksheet.Cells[row, 4].Text;   // Columna 4: Servicio
                        string membresia = worksheet.Cells[row, 6].Text;  // Columna 6: Membresia

                        // Crea y retorna un objeto de tipo InfCliente
                        return new ElementoCola(vehiculo, modelo, placa, servicio, membresia);
                    }
                }
            }

            // Si no se encuentra la membresía
            return null; // O puedes devolver un valor por defecto si prefieres
        }

        private InfCliente dt(string inputMembresia)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    for (int row = 1; row <= 100; row++) // Lee hasta la fila 100
                    {
                        string cellMembresia = worksheet.Cells[row, 4].Text; // Membresía en la columna 4

                        Debug.WriteLine($"Comparando Membresía: '{cellMembresia}' con '{inputMembresia}'");

                        if (cellMembresia == inputMembresia)
                        {
                            Debug.WriteLine("Membresía encontrada, obteniendo valores asociados.");

                            string destino = worksheet.Cells[row, 3].Text;
                            string nombre = worksheet.Cells[row, 1].Text;
                            string ci = worksheet.Cells[row, 5].Text;

                            // Crea y retorna un objeto de tipo InfCliente
                            return new InfCliente(nombre, ci, destino);
                        }
                    }
                }

                // Si no se encuentra la membresía, retorna null
                Debug.WriteLine("Membresía no encontrada.");
                return null;
            }
            catch (Exception ex)
            {
                // Maneja posibles errores al abrir el archivo o acceder al Excel
                Debug.WriteLine($"Error al procesar el archivo Excel: {ex.Message}");
                return null;
            }
        }




        private List<int> filasAEliminar = new List<int>();
        private void button6_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

            List<ElementoCola> listaVehiculos = new List<ElementoCola>(); // Lista donde guardamos los vehículos encontrados

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[2]; // Hoja 2 (origen)
                var rowCount = worksheet.Dimension.Rows;

                bool encontrada = false;

                // Recorrer todas las filas buscando la membresía en la columna 5
                for (int row = 2; row <= rowCount; row++) // Comienza desde la fila 2 para evitar el encabezado
                {
                    string membresia = worksheet.Cells[row, 5].Text; // Membresía en la columna 5
                    string servicio = worksheet.Cells[row, 4].Text; // Servicio en la columna 4
                    string vehiculo = worksheet.Cells[row, 1].Text; // Vehículo en la columna 1
                    string modelo = worksheet.Cells[row, 2].Text; // Modelo en la columna 2
                    string placa = worksheet.Cells[row, 3].Text; // Placa en la columna 3

                    int monto = 0;

                    // Si encontramos la membresía buscada, obtenemos la información del vehículo
                    if (membresia == textBox10.Text)
                    {
                        encontrada = true;

                        // Calcular el monto en base al servicio y el tipo de vehículo
                        if (servicio == "Aseo-Aspirado-Secado")
                        {
                            if (vehiculo == "Camioneta") monto = 20; else monto = 10;
                        }
                        else if (servicio == "Cambio-Aceite")
                        {
                            if (vehiculo == "Camioneta") monto = 30; else monto = 20;
                        }
                        else if (servicio == "Balanceo")
                        {
                            if (vehiculo == "Camioneta") monto = 40; else monto = 30;
                        }

                        // Crear un nuevo objeto ElementoCola con la información de este vehículo
                        ElementoCola elemento = new ElementoCola(vehiculo, modelo, placa, membresia, servicio);

                        // Asignar el monto a la PilaOpcional (por ejemplo, puedes agregarlo aquí)
                        elemento.PilaOpcional.Push(monto);

                        // Agregar el objeto a la lista
                        listaVehiculos.Add(elemento);
                        worksheet.DeleteRow(row);
                    }
                }

                // Mostrar el mensaje adecuado
                MessageBox.Show(encontrada ? "Servicio Pagado Exitosamente" : "Servicio No Encontrado");

                // Guardar los cambios realizados en el archivo Excel
                package.Save();
            }

            // Aquí ya tienes la lista de vehículos con sus montos guardados en listaVehiculos.
            // Ahora vamos a llamar a MsgUtil.EnviarFactura y pasamos toda la lista.

            if (listaVehiculos.Count > 0)
            {
                // Pasamos la lista completa a MsgUtil.EnviarFactura
                montototal = MsgUtil.EnviarFactura(bm(textBox10.Text), dt(textBox10.Text), listaVehiculos);

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[2]; // Hoja 2 (origen)

                    // Eliminar las filas guardadas en la lista filasAEliminar
                    foreach (int fila in filasAEliminar.OrderByDescending(f => f)) // Ordenar en orden descendente para evitar problemas al eliminar
                    {
                        worksheet.DeleteRow(fila);
                    }

                    // Guardar los cambios realizados en el archivo Excel
                    package.Save();

                    // Limpiar la lista de filas a eliminar después de hacer la eliminación
                    filasAEliminar.Clear();

                    // Mostrar mensaje de éxito
                    MessageBox.Show("Filas eliminadas exitosamente");
                }

                label3.Text = "Monto Total a Pagar: " + montototal.ToString();

                panel14.Hide();
                panel3.Show();
            }
            else
            {
                MessageBox.Show("No se encontraron vehículos para la membresía proporcionada.");
            }
        }




        private void Pago_Load(object sender, EventArgs e)
        {
            panel1.Show();
            panel14.Hide();
            panel2.Hide();
            panel3.Hide();
            panel5.Hide();
        }

        private bool deuda = false;
        private bool total = false;
        public int montototal = -1;

        private bool s4 = false;
        private bool s2 = false;

        private void button1_Click(object sender, EventArgs e)
        {
            total = true;
            panel14.Show();
            panel1.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            deuda = true;
            panel2.Show();
            panel1.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Pago Completo Realizado");
            MessageBox.Show("Proceso Completado Exitosamente");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel3.Hide();
            panel5.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            s4 = true;
            MessageBox.Show("Tipo de Cuota Seleccionada: 4 Semanas");
            MessageBox.Show("Proceso Completado Exitosamente");
            exceldeuda();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            s2 = true;
            MessageBox.Show("Tipo de Cuota Seleccionada: 2 semanas");
            MessageBox.Show("Proceso Completado Exitosamente");
            exceldeuda();
        }

        private void exceldeuda()
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                // Acceder a la hoja 3
                ExcelWorksheet sheet3 = package.Workbook.Worksheets[3];

                int newRowSheet3 = sheet3.Dimension?.End.Row + 1 ?? 1; // Fila para hoja 3

                // Agregar los datos directamente a la hoja 3
                sheet3.Cells[newRowSheet3, 1].Value = textBox10.Text; // Columna 1: Membresía (textBox10.Text)

                // Columna 2: Insertar "Semanal" o "Mensual" dependiendo de las condiciones booleanas
                if (s4)
                {
                    sheet3.Cells[newRowSheet3, 2].Value = 4;
                }
                else if (s2)
                {
                    sheet3.Cells[newRowSheet3, 2].Value = 2;
                }

                // Columna 3: Insertar el monto total (montototal)
                sheet3.Cells[newRowSheet3, 3].Value = montototal;

                // Guardar los cambios en el archivo Excel
                package.Save();
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[3]; // Hoja 2 (origen)
                    var rowCount = worksheet.Dimension.Rows;
                    MessageBox.Show($"Total de filas: {rowCount - 1}");

                    // Recorrer todas las filas buscando el vehículo en la columna 1
                    bool vehiculoEncontrado = false;
                    for (int row = 2; row <= rowCount; row++) // Comienza desde la fila 2 para evitar el encabezado
                    {
                        string membresiad = worksheet.Cells[row, 1].Text; // Vehículo en la columna 1

                        // Si encontramos el vehículo que buscamos, realizar las operaciones
                        if (membresiad == textBox1.Text)
                        {
                            vehiculoEncontrado = true;

                            // Verificar y convertir el valor de la columna 2 (cuota)
                            if (worksheet.Cells[row, 2].Value != null && int.TryParse(worksheet.Cells[row, 2].Value.ToString(), out int intcuota))
                            {
                                // Verificar y convertir el valor de la columna 3 (restante)
                                if (worksheet.Cells[row, 3].Value != null && int.TryParse(worksheet.Cells[row, 3].Value.ToString(), out int intrestante))
                                {
                                    // Mostrar los valores obtenidos
                                    MessageBox.Show($"Cuota: {intcuota}, Restante: {intrestante}");

                                    // Realizar la operación entre el valor de la columna 2 y columna 3
                                    int resultado = intrestante / intcuota;

                                    // Actualizar el valor en la columna 3 con la resta del resultado
                                    worksheet.Cells[row, 3].Value = intrestante - resultado;
                                    MessageBox.Show($"Total restante a pagar: {worksheet.Cells[row, 3].Value}");


                                    // Reducir el valor de la columna 2 en 1
                                    intcuota -= 1;

                                    // Si intcuota llega a 0, borrar la fila
                                    if (intcuota == 0)
                                    {
                                        worksheet.DeleteRow(row);
                                        MessageBox.Show($"Fila eliminada porque la cuota ya fue pagada.");
                                    }
                                    else
                                    {
                                        // Actualizar el valor en la columna 2
                                        worksheet.Cells[row, 2].Value = intcuota;
                                    }

                                    break;
                                }
                                else
                                {
                                    MessageBox.Show($"Error: El valor en la columna 3 no es válido en la fila {row}.");
                                }
                            }
                            else
                            {
                                MessageBox.Show($"Error: El valor en la columna 2 no es válido en la fila {row}.");
                            }
                        }
                    }

                    if (!vehiculoEncontrado)
                    {
                        MessageBox.Show("Membresia no encontrada");
                    }

                    // Guardar los cambios realizados en el archivo de Excel
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }






        //Pago Inicio = Panel1 //
        //Pago Seleccion Completo = Panel14 //
        //Pago Seleccion Cuota = Panel2
        //Pago Metodo = Panel3 //
        //Pago Tipo de Cuota = Panel5 //
    }
}
