using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;

namespace DSIGE.Dato
{
    public class DGrupo_Detalle
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        public DGrupo_Detalle() {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Grupo_Detalle> DLista(Request_Grupo_Detalle_Select oRq)
        {
            try
            {
                List<Grupo_Detalle> lSe = new List<Grupo_Detalle>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_grupo_detalle_select", oRq.gru_id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lSe.Add(
                                    new Grupo_Detalle()
                                    {
                                        gde_id = Convert.ToInt32(iDr["gde_id"]),
                                        gde_descripcion = Convert.ToString(iDr["gde_descripcion"]),
                                        gde_abreviatura = Convert.ToString(iDr["gde_abreviatura"])
                                    }
                                );
                        }
                    }
                }

                return lSe;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
