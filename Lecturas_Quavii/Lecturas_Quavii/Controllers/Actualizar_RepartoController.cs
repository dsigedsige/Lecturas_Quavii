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
    public class Actualizar_RepartoController : Controller
    {
        //
        // GET: /Actualizar_Reparto/

        public ActionResult Inicio()
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
        public string ListadoReparto(string fecha, string tipo)
        {

            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_Negocio_Listar_Reparto_Agrupado(fecha,((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, tipo);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);
        }


        [HttpPost]
        public string ListadoReparto_detallado(string fechaAsignacion, string tipo, string cod_unidad, int id_operario, string forma)
        {

            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_Negocio_Listar_Reparto_Agrupado_detallado(fechaAsignacion, tipo, cod_unidad, id_operario, forma);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);
        }




        [HttpPost]
        public string Generando_EnvioMovil_Distribucion_Detallado(List<RepartoDetalle> ListaRepartos,string FechaAsigna, string FechaMovil, string Forma, string Tipo_recibo)
        {
            object loDatos = null;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_Negocio_Generando_EnvioMovil_Distribucion_Detallado(ListaRepartos,  FechaAsigna,  FechaMovil,((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, Forma, Tipo_recibo);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



        [HttpPost]
        public string ActualizarReparto(int id_operario, string unidad_lectura, string fechaAsignatura, int id_operario_cambiar, string tipo)
        {

            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_Negocio_Validar_Operario(id_operario_cambiar);
                if (loDatos == "no_existe")
                {
                    return _Serialize(loDatos,true);
                }
                loDatos = Objeto_Negocio.Capa_Negocio_Actualizar_Reparto_Agrupado(id_operario, unidad_lectura, fechaAsignatura, id_operario_cambiar, tipo);

                //loDatos = objeto_negocio.Capa_Negocio_Listado_Servicios(); ;

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);
        }


        [HttpPost]
        public string GetReporteReparto(string fecha)
        {
            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();

                loDatos = Objeto_Negocio.Capa_Negocio_Listar_Reparto_Reporte(fecha);

                //loDatos = objeto_negocio.Capa_Negocio_Listado_Servicios(); ;

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);
        }

        [HttpPost]
        public string Generando_Compartir_lecturas(string fecha, string cod_unidad, int  id_operario , string tipo  )
        {
            object loDatos;
            try
            {
                NCorte Objeto_Negocio = new NCorte();
                loDatos = Objeto_Negocio.Capa_negocio_Generando_Compartir_lecturas(fecha, cod_unidad, id_operario, tipo);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Serialize(loDatos, true);
        }
    



}
}
