using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;

namespace archivoexcel
{
    class ManejoClosedXml
    {

        List<persona> datosDummy;

        public ManejoClosedXml(List<persona> datosDummy)
        {
            this.datosDummy = datosDummy;

            foreach(persona a in datosDummy)
            {

            }
        }
    }
}
