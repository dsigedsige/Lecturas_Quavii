using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class NGrupo_Detalle
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        public NGrupo_Detalle() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Grupo_Detalle> NLista(Request_Grupo_Detalle_Select oRq) {
            try
            {
                return new DGrupo_Detalle().DLista(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
