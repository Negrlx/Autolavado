using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Autolavado
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            string excelFilePath = @"C:\Users\Santiago\Desktop\AutoLavado.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            // Validar que el campo de búsqueda no esté vacío
            if (string.IsNullOrWhiteSpace(textBox1.Text) || textBox1.Text.Length != 5)
            {
                MessageBox.Show("Ingrese un Codigo Valido");
                return;
            }

            string codigoBusqueda = textBox1.Text; // Lo que se buscará en la tercera columna (ISBN)

            // Llamar al método estático para verificar si el código existe en la tercera columna
            bool encontrado = ExcelUtils.BuscarCodigoEnColumna(excelFilePath, codigoBusqueda);

            if (encontrado)
            {
                Inicio sexo = new Inicio();
                this.Hide();
                sexo.Show();
                MessageBox.Show("BIENVENIDO");

            }
            else
            {
                MessageBox.Show("Codigo no encontrado");
            }
        

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
