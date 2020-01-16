using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
    public class NTecnico
    {
        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 17-09-2015
        /// </summary>
        /// <returns></returns>
        public List<Operario> NLista(Request_Operario_Empresa_Local_Opcion oRq)
        {
            try
            {
                return new DTecnico().DLista(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
