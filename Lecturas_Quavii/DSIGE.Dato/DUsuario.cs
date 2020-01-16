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
    public class DUsuario
    {

        public DUsuario()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 27-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Usuario DRecuperaClave(Request_Usuario_Recupera_Clave oRq)
        {
            try
            {
                Usuario lOb = new Usuario();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_recuperacion_clave_usuario", oRq.usu_email))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.usu_nom = Convert.ToString(iDr["usu_nom"]);
                            lOb.usu_clave = Convert.ToString(iDr["usu_clave"]);
                            lOb.usu_email = Convert.ToString(iDr["usu_email"]);
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

        public List<Usuario> getUsuarios()
        {
            try
            {
                List<Usuario> user = new List<Usuario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getUsuarios"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            user.Add(
                                    new Usuario()
                                    {
                                        usu_id = Convert.ToInt32(iDr["id_Usuario"]),
                                        usu_usuario = Convert.ToString(iDr["login_usuario"]),
                                        usu_nom = Convert.ToString(iDr["nombres_usuario"]) + " " + Convert.ToString(iDr["apellidos_usuario"]),
                                        usu_cargo = Convert.ToString(iDr["nombre_cargo"]),
                                        usu_estado = Convert.ToString(iDr["estado"]),
                                        usu_check = Convert.ToString(iDr["checked"]),
                                    }
                                );
                        }
                    }
                }

                return user;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cargo_Personal> getCargo()
        {
            try
            {
                List<Cargo_Personal> cargo = new List<Cargo_Personal>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getCargo"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            cargo.Add(
                                    new Cargo_Personal()
                                    {
                                        id_Cargo = Convert.ToInt32(iDr["id_Cargo"]),
                                        nombre_cargo = Convert.ToString(iDr["nombre_cargo"]),

                                    }
                                );
                        }
                    }
                }
                return cargo;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Perfil> gePerfil()
        {
            try
            {
                List<Perfil> perfil = new List<Perfil>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getPerfil"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            perfil.Add(
                                    new Perfil()
                                    {
                                        id_Perfil = Convert.ToInt32(iDr["id_Perfil"]),
                                        nombre_perfil = Convert.ToString(iDr["nombre_perfil"]),

                                    }
                                );
                        }
                    }
                }
                return perfil;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool InsertarUsuario(Usuario objUsuario)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("sige_RegistrarUsuario",
                    objUsuario.usu_nrodoc,
                    objUsuario.usu_apellidos,
                    objUsuario.usu_nom,
                    objUsuario.usu_email,
                    objUsuario.usu_cargo,
                    objUsuario.usu_tipo,
                    objUsuario.usu_usuario,
                    objUsuario.usu_clave,
                    objUsuario.usu_idperfil,
                    objUsuario.usu_estado,
                    objUsuario.usu_usuCrea);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Usuario> ListarUsuarios()
        {
            try
            {
                List<Usuario> user = new List<Usuario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_ListarUsuarios"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            user.Add(
                                    new Usuario()
                                    {
                                        usu_id = Convert.ToInt32(iDr["id_Usuario"]),
                                        usu_apellidos = Convert.ToString(iDr["apellidos_usuario"]),
                                        usu_nom = Convert.ToString(iDr["nombres_usuario"]),
                                        usu_nrodoc = Convert.ToString(iDr["nrodoc_usuario"]),
                                        usu_email = Convert.ToString(iDr["email_usuario"]),
                                        usu_cargo = Convert.ToString(iDr["id_Cargo"]),
                                        usu_tipo = Convert.ToString(iDr["tipo_usuario"]),
                                        usu_usuario = Convert.ToString(iDr["login_usuario"]),
                                        usu_clave = Convert.ToString(iDr["contrasenia_usuario"]),
                                        usu_idperfil = Convert.ToString(iDr["id_Perfil"]),
                                        usu_estado = Convert.ToString(iDr["estado"]),

                                    }
                                );
                        }
                    }
                }

                return user;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool ActualizarUsuario(Usuario objUsuario)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_ActualizarUsuario",
                    objUsuario.usu_nrodoc,
                    objUsuario.usu_apellidos,
                    objUsuario.usu_nom,
                    objUsuario.usu_email,
                    objUsuario.usu_cargo,
                    objUsuario.usu_tipo,
                    objUsuario.usu_usuario,
                    objUsuario.usu_clave,
                    objUsuario.usu_idperfil,
                    objUsuario.usu_estado,
                    objUsuario.usu_usuEdicion,
                    objUsuario.usu_id);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool AnularUsuario(Usuario objUsuario)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_AnularUsuario", objUsuario.usu_usuEdicion, objUsuario.usu_id);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Usuario> GetDatosUsuario(int idUsuario)
        {
            try
            {
                List<Usuario> user = new List<Usuario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_DatosUsuario",idUsuario))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            user.Add(
                                    new Usuario()
                                    {
                                        usu_id = Convert.ToInt32(iDr["id_Usuario"]),
                                        usu_apellidos = Convert.ToString(iDr["apellidos_usuario"]),
                                        usu_nom = Convert.ToString(iDr["nombres_usuario"]),                                                                          
                                        usu_cargo = Convert.ToString(iDr["nombre_cargo"]),
                                        usu_perfil = Convert.ToString(iDr["nombre_perfil"]),
                                        usu_usuario = Convert.ToString(iDr["login_usuario"]),
                                        usu_clave = Convert.ToString(iDr["contrasenia_usuario"]),                                      
                                        usu_estado = Convert.ToString(iDr["estado"]),

                                    }
                                );
                        }
                    }
                }

                return user;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool CmbiarPassUsuario(Usuario objUsuario)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_CambiarPassUsuario", objUsuario.usu_clave, objUsuario.usu_id);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        /// <summary>
        /// USUARIO -SERVICIO UPDATE O INSERT
        /// </summary>
        /// <param name="idusuario"></param>
        /// <param name="idtiposervicio"></param>
        /// <param name="estadousuario"></param>
        /// <returns></returns>
        public bool Capa_Dato_Modificar_UsuarioServicio(int idusuario, int idtiposervicio, int estadousuario)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("SP_U_USUARIO_SERVICIO", idusuario, idtiposervicio, estadousuario);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="idusuario"></param>
        /// <param name="idtiposervicio"></param>
        /// <param name="estadousuario"></param>
        /// <returns></returns>
        //public bool Capa_Dato_ListarUsuarioServ(int idusuario)
        //{
        //    try
        //    {
        //        DatabaseFactory.CreateDatabase().ExecuteNonQuery("SP_L_USUARIOXSERVICIO", idusuario);
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}



        public List<Usuario_Servicio> Capa_Dato_ListarUsuarioServ(int idusuario)
        {
            try
            {
                List<Usuario_Servicio> user = new List<Usuario_Servicio>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_L_USUARIOXSERVICIO", idusuario))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            user.Add(
                                    new Usuario_Servicio()
                                    {
                                        idOperario_Servicio = Convert.ToInt32(iDr["id_Operario_Servicio"]),
                                        idOperario = Convert.ToInt32(iDr["id_Operario"]),
                                        id_TipoServicio = Convert.ToInt32(iDr["id_Servicio"]),
                                        estado = Convert.ToInt32(iDr["estado"]),
                                        nombre_tiposervicio = Convert.ToString(iDr["nombre_tiposervicio"]),
                                        descripcion = Convert.ToString(iDr["descripcion"]),
                                        color = Convert.ToString(iDr["color"]),
                                        icono = Convert.ToString(iDr["icono"]),

                                    }
                                );
                        }
                    }
                }

                return user;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
    }
}
