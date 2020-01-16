using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE.Negocio;
using DSIGE.Modelo;
using Newtonsoft.Json;


namespace DSIGE.Web.Controllers
{
    public class CambioPasswordController : Controller
    {
        //
        // GET: /CambioPassword/

        public ActionResult CambioPassword()
        {
            return View();
        }

        public static string _Serialize(object value, bool ignore = false)
        {
            var SerializerSettings = new JsonSerializerSettings()
            {
                MaxDepth = Int32.MaxValue,
                NullValueHandling = (ignore == true ? NullValueHandling.Ignore : NullValueHandling.Include)
            };

            return JsonConvert.SerializeObject(value, Formatting.Indented, SerializerSettings);
        }



        [HttpPost]
        public string Verificando_Password(string contraseña_antigua)
        {
            object loDatos;
            try
            {
                var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

                Cls_Negocio_Usuario_Cambio_Pass obj_negocio = new Cls_Negocio_Usuario_Cambio_Pass();
                loDatos = obj_negocio.Capa_Negocio_Verificando_Pass(usuario, contraseña_antigua);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string GenerandoCambioPass(string contraseña_nueva)
        {

            bool loDatos;
            try
            {
                Cls_entidad_Usuario_Cambio_Pass obj_entidad = new Cls_entidad_Usuario_Cambio_Pass();
                Cls_Negocio_Usuario_Cambio_Pass obj_negocio = new Cls_Negocio_Usuario_Cambio_Pass();

 
                obj_entidad.contrasenia_nueva = contraseña_nueva;
                obj_entidad.id_user = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;
                loDatos = obj_negocio.Capa_Negocio_Guardar_Informacion(obj_entidad);    
                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




    }
}
