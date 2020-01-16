using DSIGE.Modelo;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;

namespace DSIGE.Dato
{
   public  class Cls_Dato_Export_trabajos_lectura
    {
       string cadenaCnx = "";


       public List<Cls_Entidad_Export_trabajos_lectura> Capa_Dato_Get_ListandoLecturas(string fechaAsigna, int TipoServicio)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<Cls_Entidad_Export_trabajos_lectura> List_lectura = new List<Cls_Entidad_Export_trabajos_lectura>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_EXPORT_TRABAJOS_LECTURA_LISTAR", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                       cmd.Parameters.Add("@tipoServicio", SqlDbType.Int).Value = TipoServicio;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               Cls_Entidad_Export_trabajos_lectura obj_entidad = new Cls_Entidad_Export_trabajos_lectura(); 
                               obj_entidad.Instalacion=Fila["Instalacion"].ToString() ;
                                    obj_entidad.Equipo=Fila["Equipo"].ToString() ;
                                    obj_entidad.Aparato=Fila["Aparato"].ToString() ;
                                    obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                                    obj_entidad.Distrito=Fila["Distrito"].ToString() ;
                                    obj_entidad.Codigo_postal=Fila["Codigo_postal"].ToString() ;
                                    obj_entidad.Poblacion=Fila["Poblacion"].ToString() ;
                                    obj_entidad.Emplazamiento=Fila["Emplazamiento"].ToString() ;
                                    obj_entidad.Suplemento_emplazamiento=Fila["Suplemento_emplazamiento"].ToString() ;
                                    obj_entidad.Lectura_anterior=Fila["Lectura_anterior"].ToString() ;
                                    obj_entidad.Fecha_planificada_lectura_anterior=Fila["Fecha_planificada_lectura_anterior"].ToString() ;
                                    obj_entidad.Fecha_planificada_lectura_actual=Fila["Fecha_planificada_lectura_actual"].ToString() ;
                                    obj_entidad.Fecha_planificada_lectura_proxima=Fila["Fecha_planificada_lectura_proxima"].ToString() ;
                                    obj_entidad.Interlocutor_comercial=Fila["Interlocutor_comercial"].ToString() ;
                                    obj_entidad.Cuenta_contrato=Fila["Cuenta_contrato"].ToString() ;
                                    obj_entidad.Tipo_Cliente=Fila["Tipo_Cliente"].ToString() ;
                                    obj_entidad.Categoria=Fila["Categoria"].ToString() ; 
                                    obj_entidad.Secuencia_lectura =Fila["Secuencia_lectura"].ToString() ;
                                    obj_entidad.Unidad_lectura=Fila["Unidad_lectura"].ToString() ;
                                    obj_entidad.Numero_lecturas_estimadas_consecutivas=Fila["Numero_lecturas_estimadas_consecutivas"].ToString() ;
                                    obj_entidad.Marca_primera_lectura=Fila["Marca_primera_lectura"].ToString() ;
                                    obj_entidad.Empresa_Lectora=Fila["Empresa_Lectora"].ToString() ;
                                    obj_entidad.Nota_1_ubicacion_aparato=Fila["Nota_1_ubicacion_aparato"].ToString() ;
                                    obj_entidad.Nota_2_ubicacion_aparato=Fila["Nota_2_ubicacion_aparato"].ToString() ;
                                    obj_entidad.Tecnico=Fila["Tecnico"].ToString() ;
                                    obj_entidad.Secuencia=Fila["Secuencia"].ToString() ;
                                    obj_entidad.Grupo = Fila["Grupo"].ToString();
                               List_lectura.Add(obj_entidad);
                           }
                       }
                   }
               }
               return List_lectura;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<Cls_Entidad_Export_trabajos_lectura> Capa_Dato_Get_ListandoLecturas_Excel(string fechaAsigna, int TipoServicio)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<Cls_Entidad_Export_trabajos_lectura> List_lectura = new List<Cls_Entidad_Export_trabajos_lectura>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_EXPORT_TRABAJOS_LECTURA_EXCEL_xx", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                       cmd.Parameters.Add("@tipoServicio", SqlDbType.Int).Value = TipoServicio;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {


                               Cls_Entidad_Export_trabajos_lectura obj_entidad = new Cls_Entidad_Export_trabajos_lectura();

                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.Instalacion = Fila["Instalacion"].ToString(); 
                                obj_entidad.Aparato = Fila["Aparato"].ToString();
                                obj_entidad.Tipo_calle = Fila["Tipo_calle"].ToString().Trim();
                                obj_entidad.Nombre_Calle = Fila["Nombre_Calle"].ToString().Trim();
                                obj_entidad.Altura_Calle = Fila["Altura_Calle"].ToString().Trim();
                                obj_entidad.Numero_Edificio = Fila["Numero_Edificio"].ToString().Trim();
                                obj_entidad.Numero_Departamento = Fila["Numero_Departamento"].ToString().Trim();

                                obj_entidad.Detalle_Construccion = Fila["Detalle_Construccion"].ToString().Trim();
                                obj_entidad.Conjunto_Vivienda = Fila["Conjunto_Vivienda"].ToString().Trim();
                                obj_entidad.Manzana_Lote = Fila["Manzana_Lote"].ToString().Trim();
                                obj_entidad.Distrito = Fila["Distrito"].ToString();             
   
                                obj_entidad.Cuenta_contrato = Fila["Cuenta_contrato"].ToString();       
                                obj_entidad.Secuencia_lectura = Fila["Secuencia_lectura"].ToString();
                                obj_entidad.Unidad_lectura = Fila["Unidad_lectura"].ToString();
                                obj_entidad.Numero_lecturas_estimadas_consecutivas = Fila["Numero_lecturas_estimadas_consecutivas"].ToString();
                                
                                obj_entidad.Empresa_Lectora = Fila["Empresa_Lectora"].ToString(); 
                                obj_entidad.Nota_2_ubicacion_aparato = Fila["Nota_2_ubicacion_aparato"].ToString();
                                obj_entidad.Tecnico = Fila["Tecnico"].ToString();
                                obj_entidad.Secuencia = Fila["Secuencia"].ToString(); 

                               List_lectura.Add(obj_entidad);
                           }
                       }
                   }
               }
               return List_lectura;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

         public DataTable Capa_Dato_Get_ListaCortesReconexion_Excel(int Local , string fechaAsigna, int TipoServicio)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_EXPORT_TRABAJOS_CORTES_RECONEXION_EXCEL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@tipoServicio", SqlDbType.Int).Value = TipoServicio;
                        cmd.Parameters.Add("@idLocal", SqlDbType.Int).Value = Local;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
                return dt_detalle;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        DataTable TablePrincipal = new DataTable();

       private void GenerarCabeceraDatatable()
       {

           TablePrincipal.Columns.Add("medidorlectura", typeof(string));
           TablePrincipal.Columns.Add("mesUno", typeof(string));
           TablePrincipal.Columns.Add("mesDos", typeof(string));
           TablePrincipal.Columns.Add("mesTres", typeof(string));
           TablePrincipal.Columns.Add("mesCuatro", typeof(string));
           TablePrincipal.Columns.Add("mesCinco", typeof(string));
           TablePrincipal.Columns.Add("consumo1", typeof(string));
           TablePrincipal.Columns.Add("consumo2", typeof(string));
           TablePrincipal.Columns.Add("consumo3", typeof(string));
           TablePrincipal.Columns.Add("consumo4", typeof(string));
           TablePrincipal.Columns.Add("promedioConsumo", typeof(decimal));
           TablePrincipal.Columns.Add("parametroMax", typeof(decimal));
           TablePrincipal.Columns.Add("parametroMin", typeof(decimal));
           TablePrincipal.Columns.Add("lecturaMax", typeof(decimal));
           TablePrincipal.Columns.Add("lecMin", typeof(decimal));
       }

       public string Capa_Dato_InsertandoUpdate_LecturaMeses(int id_user, string fechaAsignacion)
       {
           var Resultado = "";

           DataTable dt = new DataTable();
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
 
              
                   using (SqlConnection con = new SqlConnection(cadenaCnx))
                   {
                       con.Open();
                       //generando tabla
                       GenerarCabeceraDatatable();
                       // generando la estructura

                       //LLenando la tabla Temporal

                       using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_TEMPORAL_LLENADO_DATA", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;
                           cmd.Parameters.Add("@fechaAsigancion", SqlDbType.VarChar).Value = fechaAsignacion;
                           cmd.ExecuteNonQuery();
                       }

                       using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_LECTURA_TEMPORAL_CALCULO_CONSUMO", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;
                           cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = id_user;
                           cmd.ExecuteNonQuery();
                       }

                       // verificando si hay registros para editar

                       using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_ARCHIVO_EXCEL_DATOS_LECTURA", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;
                           cmd.Parameters.Add("@id_user", SqlDbType.Int).Value = id_user;
                           cmd.ExecuteNonQuery();
                       }
                       //haciendo update a las validaciones de temporal_lectura

                       using (SqlCommand cmd = new SqlCommand("SP_U_TEMPORAL_LECTURA_VALIDACION", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;

                           cmd.ExecuteNonQuery();
                       }
                       // insertando Nuevo Datos

                       using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_ARCHIVO_EXCEL_DATOS_LECTURA", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;
                           cmd.Parameters.Add("@id_user", SqlDbType.Int).Value = id_user;

                           DataTable dt_detalle = new DataTable();
                           using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                           {
                               da.Fill(dt_detalle);

                               foreach (DataRow fila in dt_detalle.Rows)
                               {
                                   DataRow row = TablePrincipal.NewRow();
                                   row["medidorlectura"] = fila["cuentaContrato"].ToString();
                                   row["mesUno"] = String.IsNullOrEmpty(fila["lecturaAnterior"].ToString()) ? 0 : fila["lecturaAnterior"];
                                   row["mesDos"] = "0";
                                   row["mesTres"] = "0";
                                   row["mesCuatro"] = "0";
                                   row["mesCinco"] = "0";
                                   row["consumo1"] = "0";
                                   row["consumo2"] = "0";
                                   row["consumo3"] = "0";
                                   row["consumo4"] = "0";
                                   row["promedioConsumo"] = (decimal)0;
                                   row["parametroMax"] = (decimal)0;
                                   row["parametroMin"] = (decimal)0;
                                   row["lecturaMax"] = (decimal)0;
                                   row["lecMin"] = (decimal)0;
                                   TablePrincipal.Rows.Add(row);
                               }
                           }

                           if (dt_detalle.Rows.Count > 0)
                           {
                               using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                               {
                                   bulkCopy.BatchSize = 500;
                                   bulkCopy.NotifyAfter = 1000;
                                   bulkCopy.DestinationTableName = "tbl_lectura_Meses";
                                   bulkCopy.WriteToServer(TablePrincipal);
                               }
                           }
                       }
                   }

                  
                   Resultado = "1|Ok";
            



               return Resultado;
           }
           catch (Exception e)
           {
               Resultado = "0|" + e.Message;
           }

           return Resultado;
       }

      }
}
