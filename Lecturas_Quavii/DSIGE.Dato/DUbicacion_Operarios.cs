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
    public class DUbicacion_Operarios
    {

        public DUbicacion_Operarios()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<UbicacionOperario> DListaOperariosGPS(string fechaGPS, int idSupervisor, int idPerfil, string lista, int id_supervisor, int id_operario_supervisor, int id_local)
        {
            try
            {
                List<UbicacionOperario> lOb = new List<UbicacionOperario>();

                string[] Servicios = lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                int total = Servicios.Length;
 
                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_UBICACION_OPERARIOS_II", fechaGPS, idSupervisor, idPerfil, Servicios[0], id_supervisor, id_operario_supervisor, id_local))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new UbicacionOperario()
                                    {
                                        id_gps = Convert.ToInt32(iDr["idGPS"]),
                                        id_ope = Convert.ToInt32(iDr["id_ope"]),
                                        ope_nombre = Convert.ToString(iDr["nom_ope"]),
                                        ope_tipo = Convert.ToString(iDr["tipo_ope"]),
                                        latitud = Convert.ToString(iDr["latitud"]),
                                        longitud = Convert.ToString(iDr["longitud"]),
                                        totAsig = Convert.ToInt32(iDr["totAsig"]),
                                        totReali = Convert.ToInt32(iDr["totReali"]),
                                        totPend = Convert.ToInt32(iDr["total_lec_pend"]),
                                        pocenAvance = Convert.ToInt32(iDr["porcen_avance"]),
                                        //fechaRegistro = Convert.ToDateTime(iDr["fecha_AgregaRegistro"]),
                                        isOnline = Convert.ToString(iDr["isOnline"]),
                                        gpsActivo = Convert.ToString(iDr["GpsActivo"]),
                                        estaBateria = Convert.ToString(iDr["estadoBateria"]),
                                        //fechaAndroid = Convert.ToString(iDr["fechaAndroid"]),
                                        fechaAndroid = Convert.ToDateTime(iDr["fechaAndroid"]).ToString("dd/MM/yyyy hh:mm:ss"),
                                        HoraAndroid = Convert.ToString(iDr["HoraAndroid"]),
                                        PlanDatos = Convert.ToString(iDr["PlanDatos"]),
                                        ModoAvion = Convert.ToString(iDr["ModoAvion"])
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


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 11-10-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Operario> DLlenaCombo(int idEmp, int idOpe, int idOpcion)
        {
            try
            {
                List<Operario> lOb = new List<Operario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_llena_combo_x_perfil", idEmp, idOpe, idOpcion))
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
