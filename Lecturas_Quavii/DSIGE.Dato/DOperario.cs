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
    public class DOperario
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        public DOperario() {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <returns></returns>
        public List<Operario> DLista()
        {
            try
            {
                List<Operario> lOb = new List<Operario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_operario_new", 0, 0, 0, "", "", "", "", "", "", "", "", "", 0, 0, 0,"",""))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new Operario()
                                    {
                                        ope_id = Convert.ToInt32(iDr["ope_id"]),
                                        ope_documento = Convert.ToString(iDr["ope_documento"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        ope_tipo_usuario = Convert.ToString(iDr["tipoUsuario"]),
                                        ope_celular = Convert.ToString(iDr["ope_celular"]),
                                        ope_estado = Convert.ToInt32(iDr["ope_estado"]),
                                        servicio  = Convert.ToString(iDr["servicio"]),
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
        /// Autor: jlucero
        /// Fecha: 2015-06-19
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Operario> DLista(Request_Operario_Empresa_Local_Opcion oRq)
        {
            try
            {
                List<Operario> lOb = new List<Operario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_operario_select", oRq.emp_id, oRq.loc_id, oRq.opcion))
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

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DInserta(Request_CRUD_Operario oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_operario_new", 0, oRq.emp_id, oRq.loc_id, oRq.ope_documento, oRq.ope_documento_tipo, oRq.ope_apellido, oRq.ope_nombre, oRq.ope_foto, oRq.ope_celular, oRq.ope_usuario, oRq.ope_contrasenia, oRq.ope_online, oRq.ope_estado, oRq.ope_crea, 1, oRq.ope_tipo_usuario, oRq.ope_email));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DActualiza(Request_CRUD_Operario oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_operario_new", oRq.ope_id, oRq.emp_id, oRq.loc_id, oRq.ope_documento, oRq.ope_documento_tipo, oRq.ope_apellido, oRq.ope_nombre, oRq.ope_foto, oRq.ope_celular, oRq.ope_usuario, oRq.ope_contrasenia, oRq.ope_online, oRq.ope_estado, oRq.ope_crea, 2, oRq.ope_tipo_usuario, oRq.ope_email));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DAnula(Request_CRUD_Operario oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_operario_new", oRq.ope_id, 0, 0, "", "", "", "", "", "", "", "", "", oRq.ope_estado, oRq.ope_crea, 3,"",""));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Operario DAuditoria(Request_CRUD_Operario oRq)
        {
            try
            {
                Operario oBj = new Operario();

                using (IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_operario_new", oRq.ope_id, 0, 0, "", "", "", "", "", "", "", "", "", 0, 0, 4,"",""))
                {
                    if (oDr != null)
                    {
                        while (oDr.Read())
                        {
                            oBj.crea_id = Convert.ToInt32(oDr["crea"]);
                            oBj.crea_nombre = Convert.ToString(oDr["crea_nombre"]);
                            oBj.crea_fecha = Convert.ToString(oDr["crea_fecha"]);
                            oBj.modifica_id = Convert.ToInt32(oDr["modifica"]);
                            oBj.modifica_nombre = Convert.ToString(oDr["modifica_nombre"]);
                            oBj.modifica_fecha = Convert.ToString(oDr["modifica_fecha"]);
                        }
                    }
                }

                return oBj;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Operario DObjeto(Request_CRUD_Operario oRq)
        {
            try
            {
                Operario oBj = new Operario();

                using (IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_operario_new", oRq.ope_id, 0, 0, "", "", "", "", "", "", "", "", "", 0, 0, 5, "",""))
                {
                    if (oDr != null)
                    {
                        while (oDr.Read())
                        {
                            oBj.ope_id = Convert.ToInt32(oDr["ope_id"]);
                            oBj.emp_id = Convert.ToInt32(oDr["emp_id"]);
                            oBj.loc_id = Convert.ToInt32(oDr["loc_id"]);
                            oBj.ope_documento = Convert.ToString(oDr["ope_documento"]);
                            oBj.ope_documento_tipo = Convert.ToString(oDr["ope_documento_tipo"]);
                            oBj.ope_apellido = Convert.ToString(oDr["ope_apellido"]);
                            oBj.ope_nombre = Convert.ToString(oDr["ope_nombre"]);
                            oBj.ope_foto = Convert.ToString(oDr["ope_foto"]);
                            oBj.ope_celular = Convert.ToString(oDr["ope_celular"]);
                            oBj.ope_usuario = Convert.ToString(oDr["ope_usuario"]);
                            oBj.ope_contrasenia = Convert.ToString(oDr["ope_contrasenia"]);
                            oBj.ope_online = Convert.ToString(oDr["ope_online"]);
                            oBj.ope_estado = Convert.ToInt32(oDr["ope_estado"]);
                            oBj.ope_tipo_usuario = Convert.ToString(oDr["tipoUsuario"]);
                            oBj.ope_email = Convert.ToString(oDr["email_operario"]);
                        }
                    }
                }

                return oBj;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}