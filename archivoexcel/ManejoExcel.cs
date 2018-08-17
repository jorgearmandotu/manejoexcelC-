using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace archivoexcel
{
    public class ManejoExcel
    {
        Microsoft.Office.Interop.Excel.Application xlApp;

        public ManejoExcel(List<persona> datos)
        {
            this.xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            
            // Select the Excel cells, in the range c1 to c7 in the worksheet.
            /*Range aRange = ws.get_Range("c1", "c7");

            if (aRange == null)
            {
                Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
            Object[] args = new Object[1];
            args[0] = 6;
            aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
            /*
            // Change the cells in the C1 to C7 range of the worksheet to the number 8.
            aRange.Value2 = 8;*/

            //selecinamos celdas de la hoja
            Range rangoCeldas = ws.get_Range("a1");
            
            
            if(rangoCeldas == null)
            {
                MessageBox.Show("no tielne la version correcta de office");
            }
            //llenamos las celdas de a1 a f1 en la hoja

            Object[] enc = new object[1];
            enc[0] = "Identificacion";
            rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

            rangoCeldas = ws.get_Range("b1");
            enc[0] = "Nombre";
            rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

            rangoCeldas = ws.get_Range("c1");
            enc[0] = "Apellido";
            rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

            int index = 1;

            rangoCeldas = ws.get_Range($"d{index}");
            enc[0] = "Direccion";
            rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

            rangoCeldas = ws.get_Range($"e{index}");
            enc[0] = "telefono";
            rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

            rangoCeldas = ws.get_Range($"f{index}");
            enc[0] = "cumpleaños";
            rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);


            foreach (persona a in datos)
            {
                index++;
                rangoCeldas = ws.get_Range($"a{index}");
                enc = new object[1];
                enc[0] = a.Cc;
                rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

                rangoCeldas = ws.get_Range($"b{index}");
                enc[0] = a.Nombre;
                rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

                rangoCeldas = ws.get_Range($"c{index}");
                enc[0] = a.Apellido;
                rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

                rangoCeldas = ws.get_Range($"d{index}");
                enc[0] = a.Direccion;
                rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);

                rangoCeldas = ws.get_Range($"e{index}");
                enc[0] = a.Telefono;
                rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);
                rangoCeldas = ws.get_Range($"f{index}");
                enc[0] = a.Cumpleaños;
                rangoCeldas.GetType().InvokeMember("value", BindingFlags.SetProperty, null, rangoCeldas, enc);
            }

            /*Object[] encabezados = new object[7];
            encabezados[0] = "identificacion";
            encabezados[1] = "Nombre";
            encabezados[2] = "Apellido";
            encabezados[3] = "direccion";
            encabezados[4] = "telefono";
            encabezados[5] = "cumpleaños";
            Object[] per = new object[1];
            per[0] = "A";
            rangoCeldas.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, rangoCeldas, per);*/

        }
    }
    
}
