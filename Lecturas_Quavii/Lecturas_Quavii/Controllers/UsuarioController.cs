using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE.Modelo;
using DSIGE.Negocio;

namespace DSIGE.Web.Controllers
{
    public class UsuarioController : Controller
    {
        //
        // GET: /Usuario/



        public ActionResult Inicio()
        {
            return View();
        }

        public ActionResult Perfil()
        {
            return View();
        }




        [HttpPost]
        public ActionResult JsonCargo()
        {
           
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUsuario().getCargo()),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult JsonPerfil()
        {

            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUsuario().getPerfil()),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult JsonListarUsuario()
        {

            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUsuario().ListarUsuarios()),
                ContentType = "application/json"
            };
        }


        [HttpGet]
        public ActionResult JsonRegistrarUsu(string nroDoc, string apellidos, string nombres, string email, string cargoUsu, string tipoUsu, string Usuario, string clave, string idPerfil, int estado)
        {

            try
            {

                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NUsuario().NregistrarUsu(
                        new Usuario()
                        {

                            usu_nrodoc = nroDoc,
                            usu_apellidos = apellidos,
                            usu_nom = nombres,
                            usu_email = email,
                            usu_cargo = cargoUsu,
                            usu_tipo = tipoUsu,
                            usu_usuario = Usuario,
                            usu_clave = clave,
                            usu_idperfil = idPerfil,
                            usu_estado = estado.ToString(),
                            usu_usuCrea = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,

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
        public ActionResult JsonActualizarUsu(string nroDoc, string apellidos, string nombres, string email, string cargoUsu, string tipoUsu, string Usuario, string clave, string idPerfil, int estado,int idUsu)
        {

            try
            {

                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NUsuario().NActualizarUsu(
                        new Usuario()
                        {

                            usu_nrodoc = nroDoc,
                            usu_apellidos = apellidos,
                            usu_nom = nombres,
                            usu_email = email,
                            usu_cargo = cargoUsu,
                            usu_tipo = tipoUsu,
                            usu_usuario = Usuario,
                            usu_clave = clave,
                            usu_idperfil = idPerfil,
                            usu_estado = estado.ToString(),
                            usu_usuEdicion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            usu_id = idUsu,
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
        public ActionResult JsonAnularUsu(int idUsu)
        {

            try
            {

                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NUsuario().NAnularUsu(
                        new Usuario()
                        {
                     
                            usu_usuEdicion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            usu_id = idUsu,
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


        [HttpPost]
        public ActionResult JsonGetDatosUsuario()
        {
            int idUsuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;                     
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUsuario().GetDatosUsuario(idUsuario)),
                ContentType = "application/json"
            };
        }


        [HttpGet]
        public ActionResult JsonCambiarPassUsuario(string pass)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NUsuario().CmbiarPassUsuario(
                        new Usuario()
                        {
                            usu_clave = pass,
                            usu_id = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
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
