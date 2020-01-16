using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Web.Script.Serialization;

namespace DSIGE.Web.Controllers
{
    public class PermisosUsuariosController : Controller
    {
        //
        // GET: /PermisosUsuarios/

        [HttpGet]
        [AllowAnonymous]
        public ActionResult Inicio()
        {


            var listUsuarios = new NUsuario().getUsuarios();

            var listPermisos0 = new NPermisos().GetListaPermisos("0");
            var listPermisos1 = new NPermisos().GetListaPermisos("1");
            var listPermisos2 = new NPermisos().GetListaPermisos("2");
            var listPermisos3 = new NPermisos().GetListaPermisos("3");
            var listPermisos5 = new NPermisos().GetListaPermisos("5");
            var listPermisos6 = new NPermisos().GetListaPermisos("25");
            var listPermisos7 = new NPermisos().GetListaPermisos("79");


            var persons = new Person();
            persons.Name = "reacher gilt";
            persons.Address = "100 East Way";
            persons.Age = 74;
            persons.Detalle = new NPermisos().GetListaPermisos("1");

            List<Person> listaPer = new List<Person>();

            listaPer.Add(persons);

            ViewBag.ListaPersona = listaPer;


            ViewBag.listUsuarios = listUsuarios;
            ViewBag.listModulos0 = listPermisos0;
            ViewBag.listModulos1 = listPermisos1;
            ViewBag.listModulos2 = listPermisos2;
            ViewBag.listModulos3 = listPermisos3;
            ViewBag.listModulos5 = listPermisos5;
            ViewBag.listModulos6 = listPermisos6;
            ViewBag.listModulos7 = listPermisos7;

            return View();
        }





        [HttpGet]
        [AllowAnonymous]

        public ActionResult JsonGetUsuarios(string con)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUsuario().getUsuarios(), true),
                ContentType = "application/json"
            };
        }

        [HttpGet]
        [AllowAnonymous]

        public ActionResult JsonGetPermisos(string con, string id)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NPermisos().GetListaPermisosxUsuario(new parametersPermisos()
                        {
                            per_con = con,
                            per_id = id
                        })
                    , true),
                ContentType = "application/json"
            };
        }

        [HttpGet]
        [AllowAnonymous]

        public ActionResult JsonGetPermisosPerfil(string con, string id)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NPermisos().GetListaPermisosxPerfil(new parametersPermisos()
                        {
                            per_con = con,
                            per_id = id
                        })
                    , true),
                ContentType = "application/json"
            };
        }

        [HttpGet]
        public ActionResult JsonInsertaAccesoOpciones(string idUsu, string idOpcion, string permisos_opciones)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPermisos().InsertaAccesoOpciones(
                        new Permisos()
                        {
                            usu_id = idUsu,
                            id_Opcion = idOpcion,
                            permisos_opciones = permisos_opciones,
                            usuario_creacion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                        }
                    )),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpGet]
        public ActionResult JsonActualizarAccesoOpciones(string permisos_opciones, string idUsu, string idopcion)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPermisos().ActualizarAccesoOpciones(
                        new Permisos()
                        {
                            permisos_opciones = permisos_opciones,
                            usuario_edicion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            usu_id = idUsu,
                            id_Opcion = idopcion,
                        }
                    )),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpGet]
        public ActionResult JsonEliminarAccesoOpciones(string idUsu, string idopcion)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPermisos().EliminarAccesoOpciones(
                        new Permisos()
                        {                            
                            usu_id = idUsu,
                            id_Opcion = idopcion,
                        }
                    )),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpGet]
        public ActionResult JsonGetPefiles()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NPermisos().getPerfiles()),
                ContentType = "application/json"
            };
        }

        [HttpGet]
        public ActionResult JsonInsertaPerfil_AccesoOpciones(string idPerfil, string idOpcion, string permisos_opciones)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPermisos().InsertaPerfil_AccesoOpciones(
                        new Permisos()
                        {
                            id_Perfil = idPerfil,
                            id_Opcion = idOpcion,
                            permisos_opciones = permisos_opciones,
                            usuario_creacion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                        }
                    )),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpGet]
        public ActionResult JsonActualizarPerfil_AccesoOpciones(string permisos_opciones, string idperfil, string idopcion)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPermisos().ActualizarPerfil_AccesoOpciones(
                        new Permisos()
                        {
                            permisos_opciones = permisos_opciones,
                            usuario_edicion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            id_Perfil = idperfil,
                            id_Opcion = idopcion,
                        }
                    )),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpGet]
        public ActionResult JsonEliminarPerfil_AccesoOpciones( string idperfil, string idopcion)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPermisos().EliminarPerfil_AccesoOpciones(
                        new Permisos()
                        {                            
                            id_Perfil = idperfil,
                            id_Opcion = idopcion,
                        }
                    )),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
