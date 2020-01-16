using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using System.Data.SqlClient;
using System.Data;

namespace DSIGE.Dato
{
   public  class Cls_Dato_Usuario_Cambio_Pass
    {
       string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
       
        //  string cadenaCnx = "Data Source = .; Initial Catalog = Cobra_Toma_Datos; Integrated Security = True";


        public List<Cls_entidad_Usuario_Cambio_Pass> Capa_Dato_Verificando_Pass(int id_usuario, string contraseña)
        {
            try
            {
                List<Cls_entidad_Usuario_Cambio_Pass> obj_List_Usuario_cambio_pass = new List<Cls_entidad_Usuario_Cambio_Pass>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_USUARIO_CAMBIO_PASS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Usuario",SqlDbType.Int).Value = id_usuario;
                        cmd.Parameters.Add("@contrasenia_usuario", SqlDbType.NVarChar).Value = contraseña;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_entidad_Usuario_Cambio_Pass Entidad = new Cls_entidad_Usuario_Cambio_Pass();

                                Entidad.contrasenia_Actual = Convert.ToString(row["contrasenia_usuario"].ToString());
                                obj_List_Usuario_cambio_pass.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Usuario_cambio_pass;
            }
            catch (Exception e)
            {

                throw e;
            }
        }


        //guardando 


        public bool Capa_Dato_Guardar_Informacion(Cls_entidad_Usuario_Cambio_Pass Entidad)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_U_USUARIO_CAMBIO_PASS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Usuario", SqlDbType.Int).Value = Entidad.id_user;
                        cmd.Parameters.Add("@contrasenia_usuario", SqlDbType.NVarChar).Value = Entidad.contrasenia_nueva;
 
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }



    }
}
