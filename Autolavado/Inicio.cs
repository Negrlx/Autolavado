using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Autolavado
{
    public partial class Inicio : Form
    {



        public Inicio()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Apartado para abrir la seccion de clientes

            Cliente sexo = new Cliente();
            sexo.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Apartado para abrir la seccion de Citas

            Cita sexo = new Cita();
            sexo.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Apartado para abrir la seccion de Consultas Informaticas

            Consulta sexo = new Consulta();
            sexo.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Apartado para abrir la seccion de Pagos y Factura

            Pago sexo = new Pago();
            sexo.Show();
        }

        public Cola listos = new Cola(10);

        private void button5_Click(object sender, EventArgs e)
        {

            string filePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx";

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[1]; // Hoja 1 (origen)
                var worksheet2 = package.Workbook.Worksheets[2]; // Hoja 2 (destino)
                var rowCount = worksheet.Dimension.Rows;

                bool caas = false, cca = false, cb = false;
                int destRow = worksheet2.Dimension.Rows + 1; // Comienza en la siguiente fila vacía en la hoja destino

                // Recorrer todas las filas de la Hoja 1 buscando los servicios
                for (int row = 2; row <= rowCount; row++) // Iterar de la primera a la última fila
                {
                    string servicio = worksheet.Cells[row, 4].Text; // Columna 4: Servicio

                    if (servicio == "Aseo-Aspirado-Secado" && !caas)
                    {
                        worksheet2.Cells[destRow, 1].Value = worksheet.Cells[row, 1].Text;
                        worksheet2.Cells[destRow, 2].Value = worksheet.Cells[row, 2].Text;
                        worksheet2.Cells[destRow, 3].Value = worksheet.Cells[row, 3].Text;
                        worksheet2.Cells[destRow, 4].Value = worksheet.Cells[row, 4].Text;
                        worksheet2.Cells[destRow, 5].Value = worksheet.Cells[row, 6].Text;

                        worksheet.DeleteRow(row);
                        caas = true;
                        destRow++;
                        rowCount--; // Reducir el número de filas al eliminar una fila
                        row--; // Retroceder para procesar la nueva fila que ocupa el lugar de la eliminada
                    }
                    else if (servicio == "Balanceo" && !cca)
                    {
                        worksheet2.Cells[destRow, 1].Value = worksheet.Cells[row, 1].Text;
                        worksheet2.Cells[destRow, 2].Value = worksheet.Cells[row, 2].Text;
                        worksheet2.Cells[destRow, 3].Value = worksheet.Cells[row, 3].Text;
                        worksheet2.Cells[destRow, 4].Value = worksheet.Cells[row, 4].Text;
                        worksheet2.Cells[destRow, 5].Value = worksheet.Cells[row, 6].Text;

                        worksheet.DeleteRow(row);
                        cca = true;
                        destRow++;
                        rowCount--; // Reducir el número de filas al eliminar una fila
                        row--; // Retroceder para procesar la nueva fila
                    }
                    else if (servicio == "Cambio-Aceite" && !cb)
                    {
                        worksheet2.Cells[destRow, 1].Value = worksheet.Cells[row, 1].Text;
                        worksheet2.Cells[destRow, 2].Value = worksheet.Cells[row, 2].Text;
                        worksheet2.Cells[destRow, 3].Value = worksheet.Cells[row, 3].Text;
                        worksheet2.Cells[destRow, 4].Value = worksheet.Cells[row, 4].Text;
                        worksheet2.Cells[destRow, 5].Value = worksheet.Cells[row, 6].Text;

                        worksheet.DeleteRow(row);
                        cb = true;
                        destRow++;
                        rowCount--; // Reducir el número de filas al eliminar una fila
                        row--; // Retroceder para procesar la nueva fila
                    }
                }

                if (!cb || !caas || !cca)
                {
                    MessageBox.Show("No hay Vehiculos por Procesar");
                }
                else
                {
                    MessageBox.Show("Vehiculos Procesados Exitosamente");
                }

                // Guardar los cambios realizados en el archivo Excel
                package.Save();
            }



        }
    }
}
