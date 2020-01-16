using DSIGE.Dato;
using DSIGE.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Negocio
{
    public class NUbicacion_Operarios
    {
        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <returns></returns>
        public List<UbicacionOperario> NListaOperariosGPS(string fechaGPS, int idSupervisor, int idPerfil, string lista, int id_supervisor, int id_operario_supervisor, int id_local)
        {
            try
            {
                return new DUbicacion_Operarios().DListaOperariosGPS(fechaGPS, idSupervisor, idPerfil, lista, id_supervisor, id_operario_supervisor, id_local);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 11-10-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Operario> NLlenaCombo(int idEmp, int idOpe, int idOpcion)
        {
            try
            {
                return new DUbicacion_Operarios().DLlenaCombo(idEmp, idOpe, idOpcion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
