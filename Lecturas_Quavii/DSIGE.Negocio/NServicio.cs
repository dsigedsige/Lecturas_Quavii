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
    public class NServicio
    {
        public NServicio() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <returns></returns>
        public List<Servicio> NLista() {
            try {
                return new DServicio().DLista();
            }
            catch (Exception e) {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-19
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Servicio> NLista(Request_Servicio_Operario oRq)
        {
            try
            {
                return new DServicio().DLista(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NInserta(Request_CRUD_Servicio oRq)
        {
            try
            {
                return new DServicio().DInserta(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NActualiza(Request_CRUD_Servicio oRq)
        {
            try
            {
                return new DServicio().DActualiza(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NAnula(Request_CRUD_Servicio oRq)
        {
            try
            {
                return new DServicio().DAnula(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Servicio NAuditoria(Request_CRUD_Servicio oRq)
        {
            try
            {
                return new DServicio().DAuditoria(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Servicio NObjeto(Request_CRUD_Servicio oRq)
        {
            try
            {
                return new DServicio().DObjeto(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
