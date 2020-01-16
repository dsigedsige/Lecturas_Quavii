using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class EventoMovilOperarioSeguimientoController : Controller
    {
        //
        // GET: /EventoMovilOperarioSeguimiento/

        public ActionResult EventoMovilOperarioSeguimiento()
        {
            return View();
        }


        public ActionResult EventoMovilOperarioSeguimiento_MAPA(int id_operario, string fecha_proceso, string Servicio_proceso)
        {

            ViewBag.Operario = id_operario;
            ViewBag.FechaProceso=fecha_proceso;
            ViewBag.Servicioproceso = Servicio_proceso;
 
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
        public string ListandoDatosCabecera_Seguimiento(string id_local, int id_operario, string fecha_ini, string fecha_fin, string lista)
        {
            object loDatos;
            try
            {
                Cls_Negocio_Evento_Movil_Operario obj_negocio = new Cls_Negocio_Evento_Movil_Operario();
                loDatos = obj_negocio.Capa_Negocio_Listando_Datos_Cabecera_Seguimiento_Operario(id_local, id_operario, fecha_ini, fecha_fin, lista);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
