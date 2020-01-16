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
    public class DSesion
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-12
        /// Descripcion: Declaracion de procedimientos para el inicio de sesion
        /// </summary>
        public DSesion()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Fecha: 2015-06-12
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Sesion DGetSesion(Request_Sesion oRq) {
            try
            {
                Sesion oSe = new Sesion();

                using (DataSet oDs = DatabaseFactory.CreateDatabase().ExecuteDataSet("dsige_sesion_usuario_empresa", oRq.usuario, oRq.contrasenia))
                {
                    if (oDs != null)
                    {
                        #region Dato de accesos y usuario

                        if (oDs.Tables[0] != null)
                        {
                            using (DataTable oDt = oDs.Tables[0])
                            {
                                #region Usuario

                                Sesion_Usuario oUs = new Sesion_Usuario();

                                foreach (var usu in (from a in oDt.AsEnumerable()
                                                     group a by new
                                                     {
                                                         aop_descripcion = a.Field<string>("aop_descripcion"),
                                                         usu_id = a.Field<Int32>("usu_id"),
                                                         usu_nombre = a.Field<string>("usu_nombre"),
                                                         idPerfil = a.Field<Int32>("id_Perfil")
                                                     } into b
                                                     select new
                                                     {
                                                         aop_descripcion = b.Key.aop_descripcion,
                                                         usu_id = b.Key.usu_id,
                                                         usu_nombre = b.Key.usu_nombre,
                                                         idPerfil = b.Key.idPerfil
                                                     })) {

                                        oUs.aop_descripcion = usu.aop_descripcion;
                                        oUs.usu_id = usu.usu_id;
                                        oUs.usu_nombre = usu.usu_nombre;
                                        oUs.idPerfil = usu.idPerfil;

                                }

                                oSe.usuario = oUs;

                                #endregion

                                #region Modulo

                                List<Sesion_Modulo> lMo = new List<Sesion_Modulo>();

                                foreach (var mod in (from a in oDt.AsEnumerable()
                                                     group a by new
                                                     {
                                                         mod_id = a.Field<Int32>("mod_id"),
                                                         mod_descripcion = a.Field<string>("mod_descripcion"),
                                                         nombre_usuario = a.Field<string>("usu_nombre")

                                                     } into b
                                                     select new
                                                     {
                                                         mod_id = b.Key.mod_id,
                                                         mod_descripcion = b.Key.mod_descripcion,
                                                         nombre_usuario = b.Key.nombre_usuario
                                                     })) {
                                    Sesion_Modulo oMo = new Sesion_Modulo();
                                    List<Sesion_Item> lIt = new List<Sesion_Item>();

                                    foreach (var ite in (from a in oDt.AsEnumerable()
                                                         where a.Field<Int32>("mod_id") == mod.mod_id
                                                         select new
                                                         {
                                                             dop_id = a.Field<Int32>("dop_id"),
                                                             dop_descripcion = a.Field<string>("dop_descripcion"),
                                                             dop_url = a.Field<string>("dop_url"),
                                                             urlImagen = a.Field<string>("dop_imagen")
                                                         })) {
                                        Sesion_Item oIt = new Sesion_Item();

                                        oIt.dop_id = ite.dop_id;
                                        oIt.dop_descripcion = ite.dop_descripcion;
                                        oIt.dop_url = ite.dop_url;
                                        oIt.urlImagen = ite.urlImagen;

                                        lIt.Add(oIt);
                                    }

                                    oMo.mod_id = mod.mod_id;
                                    oMo.mod_descripcion = mod.mod_descripcion;
                                    oMo.urlImagen = mod.nombre_usuario;
                                    oMo.nombre_usuario = mod.nombre_usuario;

                                    if (mod.mod_descripcion == "MANTENIMIENTO")
                                    {
                                        oMo.urlImagen = "Menu_Mante.png";
                                    }
                                    else if (mod.mod_descripcion == "REPORTES")
                                    {
                                        oMo.urlImagen = "Menu_Reporte.png";
                                    }
                                    else if (mod.mod_descripcion == "PROCESOS")
                                    {
                                        oMo.urlImagen = "Menu_Procesos.png";
                                    }
                                    else if (mod.mod_descripcion == "NO CORTAR")
                                    {
                                        oMo.urlImagen = "Menu_Nocortar.png";
                                        oMo.dop_url = "/Industria/quavii/NoCortar/Index";
                                    }
                                    else if (mod.mod_descripcion == "SEGURIDAD")
                                    {
                                        oMo.urlImagen = "Menu_Segu.png";
                                    }


                                    oMo.item = lIt;

                                    lMo.Add(oMo);
                                }

                                oSe.modulo = lMo;

                                #endregion
                            }
                        }

                        #endregion

                        #region Dato de empresa

                        if (oDs.Tables[1] != null)
                        {
                            using (DataTable oDt = oDs.Tables[1])
                            {
                                Sesion_Empresa oEm = new Sesion_Empresa();

                                foreach (var emp in (from a in oDt.AsEnumerable()
                                                     group a by new
                                                     {
                                                         emp_id = a.Field<Int32>("emp_id"),
                                                         emp_descripcion = a.Field<string>("emp_descripcion"),
                                                         emp_razon_social = a.Field<string>("emp_razon_social"),
                                                         emp_imagen = a.Field<string>("emp_imagen")
                                                     } into b
                                                     select new
                                                     {
                                                         emp_id = b.Key.emp_id,
                                                         emp_descripcion = b.Key.emp_descripcion,
                                                         emp_razon_social = b.Key.emp_razon_social,
                                                         emp_imagen = b.Key.emp_imagen
                                                     })) {
                                    oEm.emp_id = emp.emp_id;
                                    oEm.emp_descripcion = emp.emp_descripcion;
                                    oEm.emp_razon_social = emp.emp_razon_social;
                                    oEm.emp_imagen = emp.emp_imagen;
                                }

                                oSe.empresa = oEm;
                            }
                        }

                        #endregion
                    }
                }

                return oSe;
            }
            catch (Exception e) {
                throw e;
            }
        }
    }
}