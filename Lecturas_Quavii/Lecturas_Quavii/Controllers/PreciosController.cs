using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Data.OleDb;
using System.Data;


namespace DSIGE.Web.Controllers
{
    public class PreciosController : Controller
    {
        //
        // GET: /Precios/

        public ActionResult Inicio()
        {
            return View();
        }

        [HttpPost]
        public ActionResult JsonListarLocales()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NimportarDiasTrabajo().getLocales()),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult JsonListarPrecios()
        {

            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NPrecios().ListarPrecios()),
                ContentType = "application/json"
            };
        }

        [HttpGet]
        public ActionResult JsonRegistrarPrecio(string precio, string concepto, string anio, string idLocal, string estado)
        {

            try
            {

                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPrecios().RegistrarPrecios(
                        new Precios()
                        {

                           Precio  = precio,
                            Concepto = concepto,
                            anio = anio,
                            id_Local= idLocal,
                            estado = estado,                          
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
        public ActionResult JsonActualizarPrecio(string precio, string concepto, string anio, string idLocal, string estado,int idPrecio)
        {
           try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NPrecios().ActualizarPrecios(
                        new Precios()
                        {
                            Precio = precio,
                            Concepto = concepto,
                            anio = anio,
                            id_Local = idLocal,
                            estado = estado,
                            usuario_edicion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            id_Precio = idPrecio,
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
