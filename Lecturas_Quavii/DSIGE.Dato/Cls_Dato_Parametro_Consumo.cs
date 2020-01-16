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
   public  class Cls_Dato_Parametro_Consumo
    {

 
       public Cls_Dato_Parametro_Consumo()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

       public List<Cls_Entidad_Parametro_Consumo> Capa_Dato_Get_Listado_Parametros()
       {
           try
           {
               List<Cls_Entidad_Parametro_Consumo> obj_List_Parametros_Consumo = new List<Cls_Entidad_Parametro_Consumo>();

               using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_PARAMETRO_CONSUMO"))
               {
                   if (iDr != null)
                   {
                       while (iDr.Read())
                       {
                           obj_List_Parametros_Consumo.Add(
                                   new Cls_Entidad_Parametro_Consumo()
                                   {
                                   

                               id_parametro = Convert.ToInt32(iDr["id_parametro"].ToString()),
                               nombre = Convert.ToString(iDr["nombre"].ToString()),
                               valor =Convert.ToDouble( iDr["valor"].ToString()),
                               estado =Convert.ToInt16 (iDr["estado"].ToString())
                                   }
                               );
                       }
                   }
               }

               return obj_List_Parametros_Consumo;
           }
           catch (Exception e)
           {
               throw e;
           }
       }





       //public List<Cls_Entidad_Parametro_Consumo > Capa_Dato_Get_Listado_Parametros()
       //{
       //    try
       //    {

       //        List<Cls_Entidad_Parametro_Consumo> obj_List_Parametros_Consumo = new List<Cls_Entidad_Parametro_Consumo>();

       //        using (SqlConnection cn = new SqlConnection(cadenaCnx))
       //        {
       //            cn.Open();
       //            using (SqlCommand cmd = new SqlCommand("SP_S_PARAMETRO_CONSUMO", cn))
       //            {
       //                cmd.CommandTimeout = 0;
       //                cmd.CommandType = CommandType.StoredProcedure;
     
       //                DataTable dt_detalle = new DataTable();

       //                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
       //                {
       //                    da.Fill(dt_detalle);

       //                    foreach (DataRow row in dt_detalle.Rows)
       //                    {
       //                        Cls_Entidad_Parametro_Consumo Entidad = new Cls_Entidad_Parametro_Consumo();

       //                        Entidad.id_parametro = Convert.ToInt32(row["id_parametro"].ToString());
       //                        Entidad.nombre = row["nombre"].ToString();
       //                        Entidad.valor =Convert.ToDouble( row["valor"].ToString());
       //                        Entidad.estado =Convert.ToInt16 (row["estado"].ToString());

       //                        obj_List_Parametros_Consumo.Add(Entidad);
       //                    }
       //                }
       //            }
       //        }

       //        return obj_List_Parametros_Consumo;
       //    }
       //    catch (Exception e )
       //    {
       //        throw e ;
       //    }
       //}

       //guardando 
       public bool Capa_Dato_Guardar_Informacion(Cls_Entidad_Parametro_Consumo  Entidad)
       {
           try
           {
              var xx=  Convert.ToBoolean(DatabaseFactory.CreateDatabase().ExecuteScalar("SP_I_PARAMETRO_CONSUMO", Entidad.nombre, Entidad.valor, Entidad.estado, Entidad.usuario_creacion));
               
               return true;
           }
           catch (Exception e)
           {
               throw e;
           }
       }





       //public bool Capa_Dato_Guardar_Informacion(Cls_Entidad_Parametro_Consumo  Entidad)
       //{
       //    try
       //    {
       //        using (SqlConnection cn = new SqlConnection(cadenaCnx))
       //        {
       //            cn.Open();

       //            using (SqlCommand cmd = new SqlCommand("SP_I_PARAMETRO_CONSUMO", cn))
       //            {
       //                cmd.CommandTimeout = 0;
       //                cmd.CommandType = CommandType.StoredProcedure;            
       //                cmd.Parameters.Add("@nombre", SqlDbType.VarChar).Value = Entidad.nombre;
       //                cmd.Parameters.Add("@valor", SqlDbType.Decimal).Value = Entidad.valor;
       //                cmd.Parameters.Add("@estado", SqlDbType.Int).Value = Entidad.estado;
       //                cmd.Parameters.Add("@usuario_creacion", SqlDbType.Int).Value = Entidad.usuario_creacion ;
       //                cmd.ExecuteNonQuery();
       //                return true;
       //            }
       //        }
       //    }
       //    catch (Exception e)
       //    {

       //       throw e ;
       //    }
       //}





       public bool Capa_Dato_Modificar_Informacion(Cls_Entidad_Parametro_Consumo Entidad)
       {
           try
           {
               Convert.ToBoolean(DatabaseFactory.CreateDatabase().ExecuteScalar("SP_U_PARAMETRO_CONSUMO", Entidad.id_parametro, Entidad.nombre, Entidad.valor, Entidad.estado, Entidad.usuario_edicion));

               return true;
           }
           catch (Exception e)
           {
               throw e;
           }
       }






       ///Modificando

       //public bool Capa_Dato_Modificar_Informacion(Cls_Entidad_Parametro_Consumo Entidad)
       //{
       //    try
       //    {
       //        using (SqlConnection cn = new SqlConnection(cadenaCnx))
       //        {
       //            cn.Open();

       //            using (SqlCommand cmd = new SqlCommand("SP_U_PARAMETRO_CONSUMO", cn))
       //            {
       //                cmd.CommandTimeout = 0;
       //                cmd.CommandType = CommandType.StoredProcedure;

       //                cmd.Parameters.Add("@id_parametro", SqlDbType.Int).Value = Entidad.id_parametro;
       //                cmd.Parameters.Add("@nombre", SqlDbType.VarChar).Value = Entidad.nombre;
       //                cmd.Parameters.Add("@valor", SqlDbType.Decimal).Value = Entidad.valor;
       //                cmd.Parameters.Add("@estado", SqlDbType.Int).Value = Entidad.estado;
       //                cmd.Parameters.Add("@usuario_edicion", SqlDbType.Int).Value = Entidad.usuario_edicion;                        
       //                cmd.ExecuteNonQuery();
       //                return true;
       //            }
       //        }
       //    }
       //    catch (Exception e)
       //    {

       //        throw e;
       //    }
       //}

       public bool Capa_Dato_Anular_Informacion(Cls_Entidad_Parametro_Consumo Entidad)
       {
           try
           {
               Convert.ToBoolean(DatabaseFactory.CreateDatabase().ExecuteScalar("SP_A_PARAMETRO_CONSUMO", Entidad.id_parametro,Entidad.estado, Entidad.usuario_edicion));
               return true;
           }
           catch (Exception e)
           {
               throw e;
           }
       }


       ///ANULANDO

       //public bool Capa_Dato_Anular_Informacion(Cls_Entidad_Parametro_Consumo Entidad)
       //{
       //    try
       //    {
       //        using (SqlConnection cn = new SqlConnection(cadenaCnx))
       //        {
       //            cn.Open();

       //            using (SqlCommand cmd = new SqlCommand("SP_A_PARAMETRO_CONSUMO", cn))
       //            {
       //                cmd.CommandTimeout = 0;
       //                cmd.CommandType = CommandType.StoredProcedure;

       //                cmd.Parameters.Add("@id_parametro", SqlDbType.Int).Value = Entidad.id_parametro;
       //                cmd.Parameters.Add("@estado", SqlDbType.Int).Value = Entidad.estado;
       //                cmd.Parameters.Add("@usuario_edicion", SqlDbType.Int).Value = Entidad.usuario_edicion;  
       //                cmd.ExecuteNonQuery();
       //                return true;
       //            }
       //        }
       //    }
       //    catch (Exception e)
       //    {

       //        throw e;
       //    }
       //}
       public List<Cls_Entidad_Parametro_Consumo> Capa_Dato_Get_Listado_Parametros_Auditoria(int id_parametro)
       {
           try
           {

               List<Cls_Entidad_Parametro_Consumo> obj_List_Parametros_Consumo = new List<Cls_Entidad_Parametro_Consumo>();
             

               using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_PARAMETRO_CONSUMO_AUDITORIA",id_parametro))
               {
                   if (iDr != null)
                   {
                       while (iDr.Read())
                       {
                           obj_List_Parametros_Consumo.Add(
                                   new Cls_Entidad_Parametro_Consumo()
                                   {
                        

                                        usuario_creacion = Convert.ToInt32(iDr["usuario_creacion"].ToString()),
                                        nombre_usu_crea = iDr["nombre_usu_crea"].ToString(),
                                        fecha_creacion = iDr["fecha_creacion"].ToString(),
                                        usuario_edicion = Convert.ToInt16(iDr["usuario_edicion"].ToString()),
                                        nombre_usu_modi = iDr["nombre_usu_modi"].ToString(),
                                        fecha_edicion = iDr["fecha_edicion"].ToString()
                                   }
                               );
                       }
                   }
               }

               return obj_List_Parametros_Consumo;
           }
           catch (Exception e)
           {
               throw e;
           }
       }
       ///auditoria
       ///
       //public List<Cls_Entidad_Parametro_Consumo> Capa_Dato_Get_Listado_Parametros_Auditoria(int id_parametro)
       //{
       //    try
       //    {

       //        List<Cls_Entidad_Parametro_Consumo> obj_List_Parametros_Consumo = new List<Cls_Entidad_Parametro_Consumo>();

       //        using (SqlConnection cn = new SqlConnection(cadenaCnx))
       //        {
       //            cn.Open();
       //            using (SqlCommand cmd = new SqlCommand("SP_S_PARAMETRO_CONSUMO_AUDITORIA", cn))
       //            {
       //                cmd.CommandTimeout = 0;
       //                cmd.CommandType = CommandType.StoredProcedure;
       //                cmd.Parameters.Add("@id_parametro", SqlDbType.Int).Value = id_parametro;
       //                cmd.ExecuteNonQuery();

       //                DataTable dt_detalle = new DataTable();

       //                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
       //                {
       //                    da.Fill(dt_detalle);

       //                    foreach (DataRow row in dt_detalle.Rows)
       //                    {
       //                        Cls_Entidad_Parametro_Consumo Entidad = new Cls_Entidad_Parametro_Consumo();
       //                        Entidad.usuario_creacion = Convert.ToInt32(row["usuario_creacion"].ToString());
       //                        Entidad.nombre_usu_crea = row["nombre_usu_crea"].ToString();
       //                        Entidad.fecha_creacion = row["fecha_creacion"].ToString();
       //                        Entidad.usuario_edicion = Convert.ToInt16(row["usuario_edicion"].ToString());
       //                        Entidad.nombre_usu_modi = row["nombre_usu_modi"].ToString();
       //                        Entidad.fecha_edicion = row["fecha_edicion"].ToString();
       //                        obj_List_Parametros_Consumo.Add(Entidad);
       //                    }
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






    }
}
