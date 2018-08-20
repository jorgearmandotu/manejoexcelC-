using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;

namespace archivoexcel
{
    public class ManejoClosedXml
    {

        List<persona> datosDummy;

        public ManejoClosedXml(List<persona> datosDummy)
        {
            this.datosDummy = datosDummy;
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Contacts");
           //  Title
            ws.Cell("B1").Value = "Contactos";

            // First Names
            ws.Cell("a2").Value = "identificacion";
            ws.Cell("b2").Value = "nombre";
            ws.Cell("c2").Value = "apellido";
            ws.Cell("d2").SetValue("cumpleaños"); // Another way to set the value

            // Last Names
            ws.Cell("e2").Value = "direccion";
            ws.Cell("f2").Value = "telefono";
            ws.Cell("C5").Value = "Rearden";
            ws.Cell("C6").SetValue("Taggart"); // Another way to set the value
            
            int fila = 3;
            foreach (persona a in datosDummy)
            {
                ws.Cell("a" + fila).Value = a.Cc;
                ws.Cell("b" + fila).Value = a.Nombre;
                ws.Cell("c" + fila).Value = a.Apellido;
                ws.Cell("d" + fila).Value = a.Cumpleaños;
                ws.Cell("e" + fila).Value = a.Direccion;
                ws.Cell("f" + fila).Value = a.Telefono;
                fila++;
            }

            wb.SaveAs("archivoClosedXml.xlsx");
        }
    }
}
