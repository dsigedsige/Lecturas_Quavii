using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using Microsoft.Practices.EnterpriseLibrary.Data;

namespace DSIGE.Dato
{
    public class DTecnico
    {
        public DTecnico() {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 17-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Operario> DLista(Request_Operario_Empresa_Local_Opcion oRq)
        {
            try
            {
                List<Operario> lOb = new List<Operario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_tecnico_asignado", oRq.loc_id, oRq.tip_serv, oRq.emp_id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new Operario()
                                    {
                                        ope_id = Convert.ToInt32(iDr["ope_id"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"])
                                    }
                                );
                        }
                    }
                }

                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
