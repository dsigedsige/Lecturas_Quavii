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
    public class NSesion
    {
        public NSesion(){ }

        /// <summary>
        /// Fecha: 2015-06-12
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Sesion NGetSesion(Request_Sesion oRq) 
        
        {
            try {
                return new DSesion().DGetSesion(oRq);
            }
            catch (Exception e) {
                throw e;
            }
        }
    }
}
