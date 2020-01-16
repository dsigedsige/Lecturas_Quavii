using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE.Modelo;
using DSIGE.Negocio ;
using Newtonsoft.Json;

namespace DSIGE.Web.Controllers
{
    public class Parametro_ConsumoController : Controller
    {
        //
        // GET: /Parametro_Consumo/

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
        public string ListandoParametrosConsumo()
        {
            object loDatos;
            try
            {
                Cls_Negocio_Parametro_Consumo  obj_negocio = new Cls_Negocio_Parametro_Consumo();
                loDatos = obj_negocio.Capa_Negocio_Get_Listado_Parametros ();

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }                     
        }


        // guardando la Data

        [HttpPost]
        public string GuardandoInformacion(int id_parametro, string nombre, double valor, int estado, string Tipo_Mant)
        {

            bool loDatos;
            try
            {
                Cls_Entidad_Parametro_Consumo obj_entidad = new Cls_Entidad_Parametro_Consumo();
                Cls_Negocio_Parametro_Consumo obj_negocio = new Cls_Negocio_Parametro_Consumo();

                var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

                	//SELECT id_parametro, nombre, valor, estado, usuario_creacion, fecha_creacion, usuario_edicion, fecha_edicion
 

                obj_entidad.id_parametro = id_parametro;
                obj_entidad.nombre = nombre;
                obj_entidad.valor = valor;
                obj_entidad.estado = estado;

                if (Tipo_Mant == "U")
                {
                    obj_entidad.usuario_edicion = usuario;
                }
                else {
                    obj_entidad.usuario_creacion = usuario;
                }


                if (Tipo_Mant == "U")
                {
                    loDatos = obj_negocio.Capa_Negocio_Modificar_Informacion(obj_entidad);
                }
                else {
                    loDatos = obj_negocio.Capa_Negocio_Guardar_Informacion (obj_entidad);
                }
                      
                   return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

 

        }



        //anulando

        [HttpPost]
        public string AnulandoRegistro(int id_parametro, int estado_anul)
        {

            bool loDatos;
            try
            {

                Cls_Entidad_Parametro_Consumo obj_entidad = new Cls_Entidad_Parametro_Consumo();
                Cls_Negocio_Parametro_Consumo obj_negocio = new Cls_Negocio_Parametro_Consumo();

                var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

                obj_entidad.id_parametro = id_parametro;
                obj_entidad.estado = estado_anul;
                obj_entidad.usuario_edicion  = usuario;

                loDatos = obj_negocio.Capa_Negocio_Anular_Informacion(obj_entidad);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
            
        }

        // AUDITORIA


        public string AuditoriaRegistro(int id_parametro)
        {

            object loDatos;
            try
            {

                Cls_Entidad_Parametro_Consumo obj_entidad = new Cls_Entidad_Parametro_Consumo();
                Cls_Negocio_Parametro_Consumo obj_negocio = new Cls_Negocio_Parametro_Consumo();

                loDatos = obj_negocio.Capa_Negocio_Get_Listado_Parametros_Auditoria(id_parametro);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }


    }
}
