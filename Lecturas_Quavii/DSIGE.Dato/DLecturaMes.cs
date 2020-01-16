using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Practices.EnterpriseLibrary.Data;
namespace DSIGE.Dato
{
   public class DLecturaMes
    {

       public DLecturaMes()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

       //public List<LecturaMes> Capa_Dato_Get_Listado_Parametros()
       //{
       //    try
       //    {
       //        List<LecturaMes> obj_List_Parametros_Consumo = new List<LecturaMes>();

       //        using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_PARAMETRO_CONSUMO"))
       //        {
       //            if (iDr != null)
       //            {
       //                while (iDr.Read())
       //                {
       //                    obj_List_Parametros_Consumo.Add(
       //                            new LecturaMes()
       //                            {
                                   

       //                        id_parametro = Convert.ToInt32(iDr["id_parametro"].ToString()),
       //                        nombre = Convert.ToString(iDr["nombre"].ToString()),
       //                        valor =Convert.ToDouble( iDr["valor"].ToString()),
       //                        estado =Convert.ToInt16 (iDr["estado"].ToString())
       //                            }
       //                        );
       //                }
       //            }
       //        }

       //        return obj_List_Parametros_Consumo;
       //    }
       //    catch (Exception e)
       //    {
       //        throw e;
       //    }
       //}



       //guardando 
       public bool Capa_Dato_Guardar_Informacion(LecturaMes  Entidad)
       {
           try
           {
               var xx = Convert.ToBoolean(DatabaseFactory.CreateDatabase().ExecuteScalar("sp_I_Lectura_Mes", Entidad.medidor_lectura, Entidad.mesUno_lectura, Entidad.mesDos_lectura, Entidad.mesTres_lectura,
                       Entidad.mesCuatro_lectura, Entidad.mesCinco_lectura, Entidad.consumo1
                       , Entidad.consumo2, Entidad.consumo3, Entidad.consumo4, Entidad.promedioConsumo, Entidad.parametroMaximo, Entidad.parametroMinimo, Entidad.lecturaMaximo, Entidad.LecturaMinimo));
               
               return true;
           }
           catch (Exception e)
           {
               throw e;
           }
       }





       //update
       public bool Capa_Dato_Modificar_Informacion(LecturaMes Entidad)
       {
           try
           {
              // if (Entidad.consumo1 != null)

               


                 bool conn= Convert.ToBoolean(DatabaseFactory.CreateDatabase().ExecuteScalar("sp_U_Lectura_Mes", 
                       Entidad.medidor_lectura, Entidad.mesUno_lectura, Entidad.mesDos_lectura,
                       Entidad.mesTres_lectura, Entidad.mesCuatro_lectura, Entidad.mesCinco_lectura
                       ,Entidad.consumo1,Entidad.consumo2,Entidad.consumo3,Entidad.consumo4,
                       Entidad.promedioConsumo,Entidad.parametroMaximo,Entidad.parametroMinimo
                     ,Entidad.lecturaMaximo,Entidad.LecturaMinimo));

               return true;
           }
           catch (Exception e)
           {
               throw e;
           
           }
 
   
       }



       //public bool Capa_Dato_Anular_Informacion(LecturaMes Entidad)
       //{
       //    try
       //    {
       //        Convert.ToBoolean(DatabaseFactory.CreateDatabase().ExecuteScalar("SP_A_PARAMETRO_CONSUMO", Entidad.id_parametro,Entidad.estado, Entidad.usuario_edicion));
       //        return true;
       //    }
       //    catch (Exception e)
       //    {
       //        throw e;
       //    }
       //}


       ///ANULANDO

       public List<LecturaMes> Capa_Dato_Get_Listado_ParametrosMesLectura(string id_parametro)
       {
           try
           {
               List<LecturaMes> obj_LecturaMes = new List<LecturaMes>();
               using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("sp_S_Lectura_Mes", id_parametro))
               {
                   if (iDr != null)
                   {
                       while (iDr.Read())
                       {
                           obj_LecturaMes.Add(
                                   new LecturaMes()
                                   {
                                       medidor_lectura = iDr["medidor_lectura"].ToString(),
                                       mesUno_lectura = iDr["mesUno_lectura"].ToString(),
                                       mesDos_lectura = iDr["mesDos_lectura"].ToString(),
                                       mesTres_lectura = iDr["mesTres_lectura"].ToString(),
                                       mesCuatro_lectura = iDr["mesCuatro_lectura"].ToString(),
                                       mesCinco_lectura = iDr["mesCinco_lectura"].ToString()
                                   }
                               );
                       }
                   }
               }

               return obj_LecturaMes;
           }
           catch (Exception e)
           {
               throw e;
           }
       }
    }
}
