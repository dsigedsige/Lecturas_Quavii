using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
  public      class NCuadro_Produccion
    {
        public List<Cuadro_Produccion> GetOperarios(string fechaInicial, string fechaFinal, int idlocal, string valor, string lista)
        {
            return new DCuadro_Produccion().GetOperarios(fechaInicial, fechaFinal, idlocal, valor,lista);
        }
        public List<Cuadro_Produccion> GetFechas(string fechaInicial, string fechaFinal, int idlocal, string valor,string lista)
        {
            return new DCuadro_Produccion().GetFechas(fechaInicial, fechaFinal, idlocal, valor, lista);
        }
        public List<Cuadro_Produccion> GetDatos(string fechaInicial, string fechaFinal, int idlocal, string valor, string lista)
        {
            return new DCuadro_Produccion().GetDatos(fechaInicial, fechaFinal, idlocal, valor,lista);
        }
    }
}
