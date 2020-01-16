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
    public class Dpermisos
    {
        public List<Permisos> GetListaPermisos(string id)
        {
            try
            {
                List<Permisos> oLs = new List<Permisos>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getPermisosOpciones", id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new Permisos()
                                    {
                                        modu_id = Convert.ToString(iDr["parentID"]),
                                        mod_des = Convert.ToString(iDr["nombreparentid"]),
                                        det_id = Convert.ToString(iDr["id_Opcion"]),
                                        det_des = Convert.ToString(iDr["nombre_opcion"]),

                                    }
                                );
                        }
                    }
                }

                return oLs;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public List<Permisos> GetListaPermisosxUsuario(parametersPermisos obj)
        {
            try
            {
                List<Permisos> oLs = new List<Permisos>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getPermisosxUsuario", obj.per_con, obj.per_id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new Permisos()
                                    {
                                        usu_id = Convert.ToString(iDr["usu_id"]),
                                        permisos_opciones = Convert.ToString(iDr["permisos_opciones"])

                                    }
                                );
                        }
                    }
                }

                return oLs;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Permisos> GetListaPermisosxPerfil(parametersPermisos obj)
        {
            try
            {
                List<Permisos> oLs = new List<Permisos>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getPermisosxPerfil", obj.per_con, obj.per_id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new Permisos()
                                    {
                                        id_Perfil = Convert.ToString(iDr["perfil_id"]),
                                        permisos_opciones = Convert.ToString(iDr["permisos_opciones"])

                                    }
                                );
                        }
                    }
                }

                return oLs;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool InsertaAccesoOpciones(Permisos objPermiso)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_RegistrarAccesoOpciones", 
                    objPermiso.usu_id, 
                    objPermiso.id_Opcion,
                    objPermiso.permisos_opciones, 
                    objPermiso.usuario_creacion);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool ActualizarAccesoOpciones(Permisos objPermiso)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_ActualizarAccesoOpciones",

                    objPermiso.permisos_opciones,
                    objPermiso.usuario_edicion,
                    objPermiso.usu_id,
                    objPermiso.id_Opcion);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool EliminarAccesoOpciones(Permisos objPermiso)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_EliminarAccesoOpciones",                  
                    objPermiso.usu_id,
                    objPermiso.id_Opcion);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Perfil> getPerfiles()
        {
            try
            {
                List<Perfil> lista = new List<Perfil>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getPerfiles"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lista.Add(
                                    new Perfil()
                                    {
                                        id_Perfil = Convert.ToInt32(iDr["id_Perfil"]),
                                        nombre_perfil = Convert.ToString(iDr["nombre_perfil"]),
                                        descripcion_perfil = Convert.ToString(iDr["descripcion_perfil"]),
                                        estado = Convert.ToString(iDr["estado"]),
                                        cheked = Convert.ToString(iDr["cheked"]),
                                    }
                                );
                        }
                    }
                }

                return lista;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool InsertaPerfil_AccesoOpciones(Permisos objPermiso)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_RegistrarPerfil_AccesoOpciones",
                    objPermiso.id_Perfil,
                    objPermiso.id_Opcion,
                    objPermiso.permisos_opciones,
                    objPermiso.usuario_creacion);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool ActualizarPerfil_AccesoOpciones(Permisos objPermiso)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_ActualizarPerfil_AccesoOpciones",

                    objPermiso.permisos_opciones,
                    objPermiso.usuario_edicion,
                    objPermiso.id_Perfil,
                    objPermiso.id_Opcion);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool EliminarPerfil_AccesoOpciones(Permisos objPermiso)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_EliminarPerfil_AccesoOpciones",                
                    objPermiso.id_Perfil,
                    objPermiso.id_Opcion);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
