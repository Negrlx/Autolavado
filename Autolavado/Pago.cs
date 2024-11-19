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

        private string bm(string inputMembresia)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                for (int row = 1; row <= 100; row++) // Lee hasta la fila 100
                {
                    string cellMembresia = worksheet.Cells[row, 4].Text; // Membresía en la columna 4

                    Debug.WriteLine($"Comparando Membresía: '{cellMembresia}' con '{inputMembresia}'");

                    if (cellMembresia == inputMembresia)
                    {
                        Debug.WriteLine("Membresía encontrada, obteniendo valor de la columna 3.");

                        // Retorna el valor de la columna 3 de la fila donde se encontró la membresía
                        return worksheet.Cells[row, 3].Text;
                    }
                }
            }

            // Si no se encuentra la membresía
            return null; // O puedes devolver un valor por defecto si prefieres
        }


        private void button6_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[2]; // Hoja 2 (origen)
                var rowCount = worksheet.Dimension.Rows;

                bool encontrada = false;

                // Recorrer todas las filas buscando la membresía en la columna 4
                for (int row = 2; row <= rowCount; row++) // Comienza desde la fila 2 para evitar el encabezado
                {
                    string membresia = worksheet.Cells[row, 5].Text;
                    string servicio = worksheet.Cells[row, 4].Text;
                    string vehiculo = worksheet.Cells[row, 1].Text;
                    string monto = "";

                    // Si encontramos la membresía buscada, eliminar la fila
                    if (membresia == textBox10.Text)
                    {
                        worksheet.DeleteRow(row);
                        rowCount--; // Reducir el número de filas
                        row--; // Volver a comprobar la misma fila después de eliminarla
                        encontrada = true;
                        
                        if (servicio == "Aseo-Aspirado-Secado")
                        {
                            if (vehiculo == "Camioneta") monto = "10"; else monto = "6";
                        }
                        if (servicio == "Cambio-Aceite")
                        {
                            if (vehiculo == "Camioneta") monto = "20"; else monto = "15";
                        }
                        if (servicio == "Balanceo")
                        {
                            if (vehiculo == "Camioneta") monto = "35"; else monto = "25";
                        }

                        MsgUtil.EnviarFactura(bm(membresia), monto);

                        break; // Si se encuentra y elimina, terminamos la búsqueda
                    }


                }

                // Mostrar el mensaje adecuado
                MessageBox.Show(encontrada ? "Servicio Pagado Exitosamente" : "Servicio No Encontrado");

                // Guardar los cambios realizados en el archivo Excel
                package.Save();
            }



        }

        private void Pago_Load(object sender, EventArgs e)
        {

        }
    }
}
