using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace Autolavado
{
    using System;
    using System.Collections.Generic;

    public class Node<T>
    {
        public T Dato { get; private set; }
        public Node<T> NextNode { get; set; }  // Cambio aquí para permitir modificar directamente el siguiente nodo

        public Node(T dato)
        {
            Dato = dato;
            NextNode = null;
        }
    }

    public class Lista<T>
    {
        public int Cant { get; set; }
        public Node<T> Head;
        public Node<T> Last;

        public Lista()
        {
            Head = Last = null;
            Cant = 0;
        }

        public bool ListaVacia() => Head == null;

        public void IngresarAlInicio(T dato)
        {
            var nuevoNodo = new Node<T>(dato);
            if (ListaVacia())
            {
                Head = Last = nuevoNodo;
            }
            else
            {
                nuevoNodo.NextNode = Head;
                Head = nuevoNodo;
            }
            Cant++;
        }

        public void IngresarAlFinal(T dato)
        {
            var nuevoNodo = new Node<T>(dato);
            if (ListaVacia())
            {
                Head = Last = nuevoNodo;
            }
            else
            {
                Last.NextNode = nuevoNodo;
                Last = nuevoNodo;
            }
            Cant++;
        }

        public T RetirarAlInicio()
        {
            if (ListaVacia())
                throw new InvalidOperationException("La lista está vacía.");

            T dato = Head.Dato;
            Head = Head.NextNode;

            if (Head == null)
            {
                Last = null;
            }

            Cant--;
            return dato;
        }

        public T RetirarAlFinal()
        {
            if (ListaVacia())
                throw new InvalidOperationException("La lista está vacía.");

            T dato = Last.Dato;

            if (Head == Last)
            {
                Head = Last = null;
            }
            else
            {
                var current = Head;
                while (current.NextNode != Last)
                {
                    current = current.NextNode;
                }
                current.NextNode = null;
                Last = current;
            }

            Cant--;
            return dato;
        }

        public Node<T> Localizar(T dato)
        {
            var current = Head;
            while (current != null)
            {
                if (EqualityComparer<T>.Default.Equals(dato, current.Dato))
                {
                    return current;
                }
                current = current.NextNode;
            }
            return null;
        }

        public bool Eliminar(T dato)
        {
            var current = Head;
            Node<T> previous = null;

            while (current != null)
            {
                if (EqualityComparer<T>.Default.Equals(dato, current.Dato))
                {
                    if (previous == null)
                    {
                        Head = current.NextNode;
                    }
                    else
                    {
                        previous.NextNode = current.NextNode;
                    }

                    if (current == Last)
                    {
                        Last = previous;
                    }

                    Cant--;
                    return true;
                }

                previous = current;
                current = current.NextNode;
            }
            return false;
        }

        public void Mostrar()
        {
            var current = Head;
            while (current != null)
            {
                Console.WriteLine(current.Dato);
                current = current.NextNode;
            }
        }
    }

    public class Pila<T>
    {
        private readonly Lista<T> _lista;

        public Pila()
        {
            _lista = new Lista<T>();
        }

        public void Push(T elemento) => _lista.IngresarAlInicio(elemento);

        public T Pop() => _lista.RetirarAlInicio();

        public T Peek()
        {
            if (_lista.ListaVacia())
                throw new InvalidOperationException("La pila está vacía.");

            return _lista.Localizar(_lista.RetirarAlInicio()).Dato;
        }

        public int Count() => _lista.Cant;

        public bool IsEmpty() => _lista.ListaVacia();

        public List<T> ToList()
        {
            List<T> result = new List<T>();
            var currentNode = _lista.Head; // Asumiendo que _lista tiene un nodo primero

            while (currentNode != null)
            {
                result.Add(currentNode.Dato);
                currentNode = currentNode.NextNode; // Suponiendo que la lista es enlazada
            }

            return result;
        }
    }

    public class ElementoCola
    {
        public string Vehiculo { get; set; }
        public string Modelo { get; set; }
        public string Placa { get; set; }
        public string Membresia { get; set; }
        public string Servicio { get; set; }
        public Pila<int> PilaOpcional { get; private set; }

        public ElementoCola(string vehiculo, string modelo, string placa, string membresia, string servicio)
        {
            Vehiculo = vehiculo;
            Modelo = modelo;
            Placa = placa;
            Membresia = membresia;
            Servicio = servicio;
            PilaOpcional = new Pila<int>();
        }

        public void AsignarPila(Pila<int> pila)
        {
            PilaOpcional = pila;
        }
    }

    public class InfCliente
    {
        public string Nombre { get; set; }
        public string Cedula { get; set; }
        public string Destino { get; set; }

        public InfCliente(string nombre, string ci, string mail)
        {
            Nombre = nombre;
            Cedula = ci;
            Destino = mail;
        }
    }

    public class Cola
    {
        private Lista<ElementoCola> _lista = new Lista<ElementoCola>();

        // Método público para insertar un elemento en la cola (FIFO)
        public void Insertar(ElementoCola elemento)
        {
            _lista.IngresarAlFinal(elemento); // Insertar siempre al final para cumplir FIFO
        }

        // Método público para retirar el primer elemento de la cola (FIFO)
        public ElementoCola Retirar()
        {
            if (EsVacia())
                throw new InvalidOperationException("La cola está vacía.");

            return _lista.RetirarAlInicio();
        }

        // Método para verificar si la cola está vacía
        public bool EsVacia() => _lista.ListaVacia();

        // Método para obtener la cantidad de elementos en la cola
        public int Cantidad() => _lista.Cant;

        // Método para obtener el primer elemento de la cola sin retirarlo
        public ElementoCola Inicio()
        {
            if (EsVacia())
                throw new InvalidOperationException("La cola está vacía.");

            return _lista.Head.Dato;  // Retorna el primer elemento sin retirarlo
        }

        // Método para buscar la posición de un elemento en la cola por su membresía
        public int BuscarPosicionMembresia(string membresia)
        {
            var current = _lista.Head;
            int posicion = 0;

            while (current != null)
            {
                if (current.Dato.Membresia == membresia)
                {
                    return posicion;
                }
                current = current.NextNode;
                posicion++;
            }

            return -1; // Si no encontramos el elemento, retornamos -1
        }

        // Método para eliminar un elemento de la cola por su membresía
        public void EliminarElementoPorMembresia(string membresia)
        {
            var current = _lista.Head;
            Node<ElementoCola> previous = null;

            while (current != null)
            {
                if (current.Dato.Membresia == membresia)
                {
                    if (previous == null)
                    {
                        _lista.Head = current.NextNode;
                    }
                    else
                    {
                        previous.NextNode = current.NextNode;
                    }

                    if (current == _lista.Last)
                    {
                        _lista.Last = previous;
                    }

                    _lista.Cant--;
                    return;
                }

                previous = current;
                current = current.NextNode;
            }
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

        public static int EnviarFactura(ElementoCola bm, InfCliente dt, List<ElementoCola> listaVehiculos)
        {
            string numeroFactura = new Random().Next(100000, 1000000).ToString();  // Número de la factura
            int totalMonto = 0;  // Inicializa el totalMonto como 0

            try
            {
                mail.To.Clear();
                mail.Subject = "¡Pago Realizado con Éxito! – Factura #" + numeroFactura;
                mail.IsBodyHtml = true; // Establece que el cuerpo será HTML

                // Obtiene la fecha actual
                string fechaHoy = DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy");

                int itemNumber = 1; // Para numerar los ítems

                // Construye las filas de los servicios
                string filasServicios = "";
                foreach (var vehiculo in listaVehiculos)
                {
                    // Acumula el monto de cada vehículo
                    decimal montoVehiculo = 0;

                    // Suma cada monto en PilaOpcional
                    foreach (var monto in vehiculo.PilaOpcional.ToList())
                    {
                        montoVehiculo += monto;
                    }

                    // Acumula el total general
                    totalMonto += (int)montoVehiculo;  // Convertimos el montoVehiculo a int y lo sumamos

                    // Construye la fila del servicio con los datos del vehículo
                    filasServicios += $@"
            <tr>
                <td style='padding: 8px; font-size: 14px; text-align: center;'>{itemNumber++}</td>
                <td style='padding: 8px; font-size: 14px; text-align: left;'>{vehiculo.Vehiculo}</td>
                <td style='padding: 8px; font-size: 14px; text-align: left;'>{vehiculo.Placa}</td>
                <td style='padding: 8px; font-size: 14px; text-align: left;'>{vehiculo.Servicio}</td>
                <td style='padding: 8px; font-size: 14px; text-align: right;'>${montoVehiculo}</td>
            </tr>";
                }

                // Cuerpo del correo en formato HTML
                mail.Body = $@"
    <html>
        <body style='font-family: Arial, sans-serif; background-color: #f4f4f9; color: #333;'>
            <div style='max-width: 600px; margin: 0 auto; padding: 20px; background-color: #ffffff; border-radius: 8px; border: 1px solid #ddd;'>
                <h2 style='color: #4CAF50; text-align: center;'>Factura de Servicio</h2>
                <p style='font-size: 16px;'><b>Fecha:</b> {fechaHoy}</p>
                <p style='font-size: 16px;'><b>Factura #:</b> {numeroFactura}</p>
                <hr style='border: 1px solid #ddd; margin: 20px 0;'>
                <p style='font-size: 16px;'><b>Cliente:</b> {dt.Nombre} <b>Cédula:</b> {dt.Cedula}</p>
                <hr style='border: 1px solid #ddd; margin: 20px 0;'>
                <table style='width: 100%; border-collapse: collapse;'>
                    <thead>
                        <tr style='background-color: #f1f1f1;'>
                            <th style='padding: 8px; font-size: 14px; text-align: left;'>Item</th>
                            <th style='padding: 8px; font-size: 14px; text-align: left;'>Vehículo</th>
                            <th style='padding: 8px; font-size: 14px; text-align: left;'>Placa</th>
                            <th style='padding: 8px; font-size: 14px; text-align: left;'>Servicio</th>
                            <th style='padding: 8px; font-size: 14px; text-align: right;'>Monto</th>
                        </tr>
                    </thead>
                    <tbody>
                        {filasServicios}
                    </tbody>
                </table>
                <hr style='border: 1px solid #ddd; margin: 20px 0;'>
                <p style='font-size: 16px; text-align: right;'><b>Total:</b> ${totalMonto}</p>
                <hr style='border: 1px solid #ddd; margin: 20px 0;'>
                <p style='font-size: 16px;'>Gracias por confiar en nosotros. Si tienes alguna pregunta, no dudes en contactarnos.</p>
                <p style='font-size: 14px; color: #777; margin-top: 30px;'>Atentamente,<br/> El Equipo de Autoprocure</p>
            </div>
        </body>
    </html>";

                mail.From = new MailAddress("autoprocurevzla@gmail.com");
                mail.To.Add(dt.Destino.Trim());

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

            return totalMonto;  // Retorna el totalMonto como un int
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