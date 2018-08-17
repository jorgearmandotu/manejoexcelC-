using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace archivoexcel
{
    public class persona
    {
        private String cc;
        private String nombre;
        private String apellido;
        private String direccion;
        private String telefono;
        private String cumpleaños;

        public persona(string cc, string nombre, string apellido, string direccion, string telefono, string cumpleaños)
        {
            this.Cc = cc;
            this.Nombre = nombre;
            this.Apellido = apellido;
            this.Direccion = direccion;
            this.Telefono = telefono;
            this.Cumpleaños = cumpleaños;
        }

        public string Cc { get => cc; set => cc = value; }
        public string Nombre { get => nombre; set => nombre = value; }
        public string Apellido { get => apellido; set => apellido = value; }
        public string Direccion { get => direccion; set => direccion = value; }
        public string Telefono { get => telefono; set => telefono = value; }
        public string Cumpleaños { get => cumpleaños; set => cumpleaños = value; }
    }
}
