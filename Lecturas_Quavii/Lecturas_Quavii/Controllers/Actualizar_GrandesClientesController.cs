using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Lecturas_Quavii.Controllers
{
    public class Actualizar_GrandesClientesController : Controller
    {
        //
        // GET: /Actualizar_GrandesClientes/

        public ActionResult grandesClientes_Index()
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
        public string get_listandoGrandesClientes_agrupado(string fecha)
        {

            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_Negocio_get_grandesClientes_Agrupado(fecha, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);
        }


        [HttpPost]
        public string Actualizar_grandesClientes(int id_operario, string distrito, string fechaAsignatura, int id_operario_cambiar)
        {
            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                var respuesta = Objeto_Negocio.Capa_Negocio_Validar_Operario(id_operario_cambiar);
                if (respuesta == "no_existe")
                {
                    return _Serialize(respuesta, true);
                }
                else {
                    loDatos = Objeto_Negocio.Capa_Negocio_Actualizar_grandesClientes_Agrupado(id_operario, distrito, fechaAsignatura, id_operario_cambiar);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Serialize(loDatos, true);
        }


        [HttpPost]
        public string get_grandesClientes_detallado(string fechaAsignacion, string distrito, int id_operario, string forma)
        {

            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_Negocio_get_grandesClientes_Agrupado_detallado(fechaAsignacion, distrito, id_operario, forma);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);
        }
               

        [HttpPost]
        public string Generando_EnvioMovil_grandesClientes_Detallado(List<GrandesClientesDetalle> ListaGrandesClientes, string FechaAsigna, string FechaMovil, string Forma)
        {
            object loDatos = null;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_Negocio_Generando_EnvioMovil_grandesClientes_Detallado(ListaGrandesClientes, FechaAsigna, FechaMovil, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, Forma);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }






    }
}
