using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;


namespace Autolavado
{

    public class Pila
    {
        private const int MAX = 4; // Tamaño máximo de la pila
        private int[] datos;    // Array para almacenar los elementos
        private int tope;           // Índice del tope de la pila
        private int cant;           // Cantidad de elementos en la pila

        public Pila()
        {
            datos = new int[MAX];
            tope = -1;
            cant = 0;
        }

        public int VerTope()
        {
            if (!PilaVacia())
            {
                return datos[tope];
            }
            else
            {
                return -1;
            }
        }

        public bool PilaLlena()
        {
            return cant == MAX;
        }

        public bool PilaVacia()
        {
            return tope == -1;
        }

        public void Push(int elemento)
        {
            if (!PilaLlena())
            {
                datos[++tope] = elemento;
                cant++;
            }
        }

        public int Pop()
        {
            if (!PilaVacia())
            {
                int elemento = datos[tope--];
                cant--;
                return elemento;
            }
            else
            {
                return -1;
            }
        }

        public int Cantidad()
        {
            return cant;
        }

        // Limpiar la pila
        public void Limpiar()
        {
            tope = -1;
            cant = 0;
        }
    }
    public class ElementoCola
    {
        public string Vehiculo { get; set; }
        public string Modelo { get; set; }
        public string Placa { get; set; }
        public string Membresia { get; set; }
        public string Servicio { get; set; }
        public Pila PilaOpcional { get; private set; }

        // Constructor para 4 strings
        public ElementoCola(string vehiculo, string modelo, string placa, string membresia, string servicio)
        {
            Vehiculo = vehiculo;
            Modelo = modelo;
            Placa = placa;
            Membresia = membresia;
            Servicio = servicio;
            PilaOpcional = null; // Inicialmente no hay pila
        }

        // Método para asignar o reemplazar la pila opcional
        public void AsignarPila(Pila pila)
        {
            PilaOpcional = pila;
        }

        // Método para imprimir los detalles del objeto
        public override string ToString()
        {
            string infoPila = PilaOpcional != null ? $"Pila (cantidad): {PilaOpcional.Cantidad()}" : "No hay pila asociada";
            return $"{Vehiculo}, {Modelo}, {Placa}, {Membresia}, {Servicio} | {infoPila}";
        }
    }
    public class Cola
    {
        int MAX;
        private ElementoCola[] datos;
        private int cant;
        private int inicio;
        private int fin;

        public Cola(int n)
        {
            MAX = n;
            datos = new ElementoCola[MAX];
            cant = 0;
            inicio = -1;
            fin = -1;
        }

        public bool Vacia()
        {
            return cant == 0;
        }

        public bool Llena()
        {
            return cant == MAX;
        }

        public void Limpiar()
        {
            cant = 0;
            inicio = -1;
            fin = -1;
        }

        public int Cantidad()
        {
            return cant;
        }

        public int Inicio()
        {
            return inicio;
        }

        public void Insertar(ElementoCola elemento)
        {
            if (!Llena())
            {
                if (inicio == -1) inicio = 0; // Configurar el inicio si es la primera inserción
                fin = (fin + 1) % MAX;
                datos[fin] = elemento;
                cant++;
            }
            
        }

        public ElementoCola Retirar()
        {
            if (!Vacia())
            {
                ElementoCola elemento = datos[inicio];
                inicio = (inicio + 1) % MAX;
                cant--;
                return elemento;
            }
            return null;
        }

        public void Eliminar(int posicion)
        {
            if (!Vacia())
            {
                if (posicion >= 0 && posicion < cant)
                {
                    int indiceEliminar = (inicio + posicion) % MAX;

                    for (int i = indiceEliminar; i != fin; i = (i + 1) % MAX)
                    {
                        int siguienteIndice = (i + 1) % MAX;
                        datos[i] = datos[siguienteIndice];
                    }

                    datos[fin] = null;
                    fin = (fin - 1 + MAX) % MAX;
                    cant--;

                    if (cant == 0)
                    {
                        inicio = -1;
                        fin = -1;
                    }
                    else if (inicio == (fin + 1) % MAX)
                    {
                        inicio = (inicio + 1) % MAX;
                    }
                }
            }
        }




        public int BuscarPosicion(string membresia)
        {
            if (!Vacia())
            {
                for (int i = 0; i < cant; i++)
                {
                    int indice = (inicio + i) % MAX;
                    if (datos[indice].Membresia == membresia)
                    {
                        return i; 
                    }
                }
            }
            return -1;
        }

    }

    public static class membresia
    {
        public static string nurandom()
        {
            var random = new Random();
            return new string(Enumerable.Repeat("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", 5)
                                        .Select(s => s[random.Next(s.Length)]).ToArray());
        }

    }   
    public static class MsgUtil
    {
        private static readonly MailMessage mail = new MailMessage();
        private static readonly SmtpClient sender = new SmtpClient("smtp.gmail.com");



        public static string CodigoVerificacion { get; private set; }

        public static void EnviarMembresia(string destino, string codigo)
        {
            try
            {
                mail.To.Clear();
                mail.Subject = "¡Bienvenido a Autoprocure! – Tu Código de Membresía";
                mail.IsBodyHtml = true; // Establece que el cuerpo será HTML

                // Cuerpo del correo en formato HTML
                mail.Body = @"
        <html>
            <body style='font-family: Arial, sans-serif; background-color: #f4f4f9; color: #333;'>
                <div style='max-width: 600px; margin: 0 auto; padding: 20px; background-color: #ffffff; border-radius: 8px; border: 1px solid #ddd;'>
                    <h1 style='color: #4CAF50; text-align: center;'>¡Bienvenido a Autoprocure!</h1>
                    <p style='font-size: 18px; text-align: center;'>Estamos encantados de tenerte como parte de nuestra comunidad.</p>
                    <p style='font-size: 16px;'>Como nuevo miembro, aquí tienes tu código exclusivo de membresía:</p>
                    <div style='text-align: center; margin: 20px 0;'>
                        <p style='font-size: 24px; font-weight: bold; color: #4CAF50; border: 2px dashed #4CAF50; display: inline-block; padding: 10px 20px; border-radius: 5px;'>
                            " + codigo + @"
                        </p>
                    </div>
                    <p style='font-size: 16px;'>Con tu membresía, tendrás acceso a beneficios exclusivos, descuentos y una experiencia de compra personalizada.</p>
                    <p style='font-size: 16px;'>Si necesitas ayuda o tienes alguna pregunta, no dudes en contactarnos. ¡Estamos aquí para ti!</p>
                    <p style='font-size: 14px; color: #777; margin-top: 30px; text-align: center;'>Atentamente,<br/><strong>El Equipo de Autoprocure</strong></p>
                </div>
            </body>
        </html>";

                mail.From = new MailAddress("autoprocurevzla@gmail.com");
                mail.To.Add(destino.Trim());

                sender.Port = 587;
                sender.UseDefaultCredentials = false;
                sender.Credentials = new System.Net.NetworkCredential("autoprocurevzla@gmail.com", "ouqp rhcd nqmz lfxd");
                sender.EnableSsl = true;

                sender.Send(mail);
                Console.WriteLine("Correo enviado exitosamente");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al enviar el correo: " + ex.Message);
            }
        }

        public static void EnviarFactura(string destino, string monto)
        {
            string numeroFactura = "522424";
            try
            {
                mail.To.Clear();
                mail.Subject = "¡Pago Realizado con Éxito! – Factura #" + numeroFactura;
                mail.IsBodyHtml = true; // Establece que el cuerpo será HTML

                // Obtiene la fecha actual
                string fechaHoy = DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy");

                // Cuerpo del correo en formato HTML
                mail.Body = @"
            <html>
                <body style='font-family: Arial, sans-serif; background-color: #f4f4f9; color: #333;'>
                    <div style='max-width: 600px; margin: 0 auto; padding: 20px; background-color: #ffffff; border-radius: 8px; border: 1px solid #ddd;'>
                        <h2 style='color: #4CAF50;'>¡Gracias por tu pago!</h2>
                        <p style='font-size: 18px;'>Querido cliente,</p>
                        <p style='font-size: 16px;'>Nos complace informarte que hemos recibido tu pago de forma exitosa. Aquí están los detalles de tu factura:</p>
                        <table style='width: 100%; border-collapse: collapse; margin-top: 20px;'>
                            <tr>
                                <td style='padding: 8px; font-size: 14px; font-weight: bold; background-color: #f1f1f1;'>Factura #</td>
                                <td style='padding: 8px; font-size: 14px; background-color: #f1f1f1;'>" + numeroFactura + @"</td>
                            </tr>
                            <tr>
                                <td style='padding: 8px; font-size: 14px; font-weight: bold; background-color: #f1f1f1;'>Monto</td>
                                <td style='padding: 8px; font-size: 14px; background-color: #f1f1f1;'>$" + monto + @"</td>
                            </tr>
                            <tr>
                                <td style='padding: 8px; font-size: 14px; font-weight: bold; background-color: #f1f1f1;'>Fecha</td>
                                <td style='padding: 8px; font-size: 14px; background-color: #f1f1f1;'>" + fechaHoy + @"</td>
                            </tr>
                        </table>
                        <p style='font-size: 16px; margin-top: 20px;'>Estamos muy agradecidos por tu confianza. Si tienes alguna pregunta, no dudes en contactarnos.</p>
                        <p style='font-size: 16px;'>¡Te deseamos un excelente día!</p>
                        <p style='font-size: 14px; color: #777; margin-top: 30px;'>Atentamente,<br/> El Equipo de Autoprocure</p>
                    </div>
                </body>
            </html>";

                mail.From = new MailAddress("autoprocurevzla@gmail.com");
                mail.To.Add(destino.Trim());

                sender.Port = 587;
                sender.UseDefaultCredentials = false;
                sender.Credentials = new System.Net.NetworkCredential("autoprocurevzla@gmail.com", "ouqp rhcd nqmz lfxd");
                sender.EnableSsl = true;

                sender.Send(mail);
                Console.WriteLine("Correo de factura enviado exitosamente");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al enviar el correo: " + ex.Message);
            }
        }



    }

    public static class ExcelUtils
    {
        // Método estático que verifica si un código se encuentra en la tercera columna
        public static bool BuscarCodigoEnColumna(string excelFilePath, string codigo)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajamos con la primera hoja

                    // Iterar sobre las filas de Excel desde la fila 2 hasta la última
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        // Verificar si el valor en la tercera columna (ISBN) coincide con el código a buscar
                        if (worksheet.Cells[row, 4 ].Text.Equals(codigo))
                        {
                            return true; // Si se encuentra el código, retornar true
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Manejar errores de lectura o archivo
                Console.WriteLine($"Error: {ex.Message}");
            }

            return false; // Si no se encuentra el código, retornar false
        }
    }


}