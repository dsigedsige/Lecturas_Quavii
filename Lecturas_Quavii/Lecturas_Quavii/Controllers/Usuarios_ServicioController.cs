using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;

namespace DSIGE.Web.Controllers
{
    public class Usuarios_ServicioController : Controller
    {
        //
        // GET: /Usuarios_Servicio/

        public ActionResult Inicio()
        {
            return View();
        }



        public ActionResult MantenimientoUsuarioServicio()
        {
            return View();
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 23-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonOperarioLocal(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NLocal().NLista(
                    new Request_Local_Operario() {
                        ope_id = __a
                    }
                )),
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 23-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult AsignaServicioL(string __a, string __b, string __c)
        {
            Int32 respuesta = new NOperario_Servicio().NAsignaServicioL(
                        new Request_Operario_Servicio_Asignacion()
                        {
                            ope_id = __a,
                            ser_id = Convert.ToInt32(__b),
                            usu_id = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            sop_estado = Convert.ToInt32(__c),
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                           //loc_id = Convert.ToInt32(__d)
                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }


        /*********************MANTENIMIENTO**********************/
        public static string _Serialize(object value, bool ignore = false)
        {
            var SerializerSettings = new JsonSerializerSettings()
            {
                MaxDepth = Int32.MaxValue,
                NullValueHandling = (ignore == true ? NullValueHandling.Ignore : NullValueHandling.Include)
            };
            return JsonConvert.SerializeObject(value, Formatting.Indented, SerializerSettings);
        }

        /// <summary>
        /// retorna una lista de usuarios
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public string ListandoServicios()
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaServicio();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

        /// <summary>
        /// devulve una Lista Usuarios
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public string ListandoUsuarios()
        {
            object loDatos;
            try
            {
                NUsuario obj_negocio = new NUsuario();
                loDatos = obj_negocio.ListarUsuarios();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

         [HttpPost]
        public string ListarUsuarioXServ(int idusuario)
        {
            object loDatos;
            try
            {
                NUsuario obj_negocio = new NUsuario();
                loDatos = obj_negocio.Capa_Negocio_Listar_UsuarioxServ(idusuario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }



        



        [HttpPost]
        public string UpdateUsuarioServicio(int idusuario, int idtiposervicio, int estadousuario)
        {

            try
            {
                object obj_list_datos = "";
              NUsuario obj_negocio = new NUsuario();
                obj_negocio.Capa_Negocio_Modificar_UsuarioServicio(idusuario,idtiposervicio,estadousuario);


                return _Serialize(obj_list_datos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }


        }


        [HttpPost]
        public string UpdateCreateUserServicio(int estado, List<int> List_cod)
        {

            
            try
            {

                object obj_list_datos = "";
                NUsuario obj_negocio = new NUsuario();
                obj_negocio.Capa_Negocio_ModificarCrear_UsuarioServicio(estado, List_cod);


                return _Serialize(obj_list_datos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

    }
}
