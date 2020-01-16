using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class seguimientoFotosController : Controller
    {
        //
        // GET: /seguimientoFotos/

        public ActionResult seguimientoFoto_index()
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
        public string ListandoLectura_Seguimiento(int id_local, int id_servicio, string fechainicial, string fechafinal)
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_ListandoLectura_Seguimiento(id_local, id_servicio, fechainicial, fechafinal, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

        [HttpPost]
        public string ListandoFotosDescargar_seguimiento(int local, int servicio, string fecha, string option)
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_ListandoFotosDescargar_seguimiento(local, servicio, fecha, option,((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string Listando_Data_Descargar_seguimiento(int local, int servicio, string fecha, string option)
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_Listando_Data_Descargar_seguimiento(local, servicio, fecha, option, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


    }
}
