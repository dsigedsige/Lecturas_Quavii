using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE.Modelo;
using DSIGE.Negocio;

namespace DSIGE.Web.Controllers
{
    public class EventoMovilOperarioController : Controller
    {
        //
        // GET: /EventoMovilOperario/

        public ActionResult EventoMovilOperario()
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
        public string ListandoLocales()
        {
            object loDatos;
            try
            {
                Cls_Negocio_Evento_Movil_Operario obj_negocio = new Cls_Negocio_Evento_Movil_Operario();
                loDatos = obj_negocio.Capa_Negocio_Locales();

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ListandoOperarioLocales(int id_local)
        {
            object loDatos;
            try
            {
                Cls_Negocio_Evento_Movil_Operario obj_negocio = new Cls_Negocio_Evento_Movil_Operario();
                loDatos = obj_negocio.Capa_Negocio_Operarios(id_local);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ListandoDatosCabecera(string id_local, string id_operario, string fecha_ini, string fecha_fin, string lista)
        {
            object loDatos;
            try
            {
                Cls_Negocio_Evento_Movil_Operario obj_negocio = new Cls_Negocio_Evento_Movil_Operario();
                loDatos = obj_negocio.Capa_Negocio_Listando_Datos_Cabecera(id_local, id_operario, fecha_ini, fecha_fin,lista);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        public string ListandoDatosOperarioDetallado( int  id_operario,  string fecha)
        {
            object loDatos;
            try
            {
                Cls_Negocio_Evento_Movil_Operario obj_negocio = new Cls_Negocio_Evento_Movil_Operario();
                loDatos = obj_negocio.Capa_Negocio_Datos_Operario_Detallado(id_operario, fecha);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




    }
}
