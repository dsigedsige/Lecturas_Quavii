using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;

namespace DSIGE.Web.Controllers
{
    public class LoginController : Controller
    {
        //
        // GET: /Login/

        /// <summary>
        /// Fecha: 2015-06-12
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [AllowAnonymous]
        public ActionResult Login()
        {
            if (Session["Session_Usuario_Acceso"] != null)
            {
                return RedirectToAction("Inicio", "Login");
            }

            return View();
        }

        /// <summary>
        /// Fecha: 2015-06-12
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Login(Request_Sesion __a) {

            if (ModelState.IsValid)
            {
                Sesion oSe = new NSesion().NGetSesion(__a);

                if (oSe == null || oSe.usuario == null || oSe.modulo.Count == 0)
                {
                    ViewBag.mensaje = "Nombre de usuario y contraseña, no son válidos.";
                }
                else {
                    ViewBag.usuario = oSe.usuario.usu_nombre.ToString().ToUpper();
                    Session["Session_Usuario_Acceso"] = oSe;
                    return RedirectToAction("Inicio", "Login");
                }
            }

            return View();
        }

        /// <summary>
        /// Fecha: 2015-06-12
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [AllowAnonymous]
        public ActionResult Logout()
        {
            Session.RemoveAll();
            return RedirectToAction("Login");
        }

        /// <summary>
        /// Fecha: 2015-06-12
        /// </summary>
        /// <returns></returns>
        public ActionResult Inicio() {
            return View();
        }
    }
}