using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using System.ComponentModel;
using System.Transactions;

namespace DSIGE.Dato
{
    public class ImportarArchivoPlano_DAO
    {

        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;


        public string Capa_Dato_GuardarArchivos(List<ImportarArchivoPlano> obj_list, int id_user, int empresa)
        {
            string resultado = "";
            string nombreArchivo = "";
            try
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(ImportarArchivoPlano));
                DataTable table = new DataTable();

                foreach (PropertyDescriptor prop in properties)
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                foreach (ImportarArchivoPlano item in obj_list)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    table.Rows.Add(row);
                }

                string nameFile = "";
                string AssingDate = "";
                int cantidadRegistros = 0;
                foreach (var row in obj_list)
                {
                    if (row.nombre_ArchivoImportado != "")
                    {
                        nameFile = row.nombre_ArchivoImportado;
                        AssingDate = row.fecha_Asignacion;
                        break;
                    }
                }

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    ///eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_IMPORTACION_ARCHIVO_TXT", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                        cmd.Parameters.Add("@nombre_archivo", SqlDbType.VarChar).Value = nameFile;
                        cmd.ExecuteNonQuery();
                    }
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {
                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "tbl_importacion_archivo_txt_T";
                        bulkCopy.WriteToServer(table);

                        //Actualizando campos 
                        string Sql = "UPDATE tbl_importacion_archivo_txt_T SET id_empresa='" + empresa + "' ,  nombre_ArchivoImportado='" + nameFile + "' ,  fecha_Asignacion='" + AssingDate + "'  , id_usuario='" + id_user + "', fecha_importacion=getdate()  WHERE id_usuario IS NULL";
                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                        //// insertando en la tabla Lecturas
                        //using (SqlCommand cmd = new SqlCommand("SP_I_ARCHIVO_TXT_INSERT_LECTURAS_II", con))
                        //{
                        //    cmd.CommandTimeout = 0;
                        //    cmd.CommandType = CommandType.StoredProcedure;
                        //    cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                        //    cmd.Parameters.Add("@nombre_archivo", SqlDbType.VarChar).Value = nameFile;

                        //    cmd.ExecuteNonQuery();
                        //}
                    }

                }
                resultado = "1|OK";

            }
            catch (Exception e)
            {
                resultado = "0|" + e.Message;
            }
            return resultado;
        }

        public string Capa_Dato_GuardarArchivos_new(List<ImportarArchivoPlano_new> obj_list, int id_user, int empresa)
        {
            string resultado = "";
            try
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(ImportarArchivoPlano_new));
                DataTable table = new DataTable();

                foreach (PropertyDescriptor prop in properties)
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                foreach (ImportarArchivoPlano_new item in obj_list)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    table.Rows.Add(row);
                }

                string nameFile = "";
                string AssingDate = "";
 
                foreach (var row in obj_list)
                {
                    if (row.nombre_ArchivoImportado != "")
                    {
                        nameFile = row.nombre_ArchivoImportado;
                        AssingDate = row.fecha_Asignacion;
                        break;
                    }
                }

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    ///eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_IMPORTACION_ARCHIVO_TXT_NEW", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                        cmd.Parameters.Add("@nombre_archivo", SqlDbType.VarChar).Value = nameFile;
                        cmd.ExecuteNonQuery();
                    }
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {
                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "tbl_importacion_archivo_txt_T_new";
                        bulkCopy.WriteToServer(table);

                        //Actualizando campos 
                        string Sql = "UPDATE tbl_importacion_archivo_txt_T_new SET id_empresa='" + empresa + "' ,  nombre_ArchivoImportado='" + nameFile + "' ,  fecha_Asignacion='" + AssingDate + "'  , id_usuario='" + id_user + "', fecha_importacion=getdate()  WHERE id_usuario IS NULL";
                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                        // insertando en la tabla Lecturas
                        using (SqlCommand cmd = new SqlCommand("SP_I_ARCHIVO_TXT_INSERT_LECTURAS_NEW_FORMAT", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                            cmd.Parameters.Add("@nombre_archivo", SqlDbType.VarChar).Value = nameFile;

                            cmd.ExecuteNonQuery();
                        }
                    }

                }
                resultado = "1|OK";

            }
            catch (Exception e)
            {
                resultado = "0|" + e.Message;
            }
            return resultado;
        }


        public List<ImportarArchivoPlano> Capa_Dato_Get_ListaArchivosAlmacenados(string fechaAsignacion, int TipoServicio)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ImportarArchivoPlano> ListArchivos = new List<ImportarArchivoPlano>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ARCHIVO_PLANO_VERIFICAR_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion_Lectura", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@tipoServicio", SqlDbType.Int).Value = TipoServicio;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                ImportarArchivoPlano obj_entidad = new ImportarArchivoPlano();
                                obj_entidad.nombre_ArchivoImportado = Fila["archivoImportacion_lectura"].ToString();
                                ListArchivos.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListArchivos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


 

    }
}
