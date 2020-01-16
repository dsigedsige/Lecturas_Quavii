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
    public class Dias_TrabajoController : Controller
    {
        //
        // GET: /Parametro_Consumo/

        public ActionResult Dias_Trabajo()
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
        public string ListandoDiasTrabajo()
        {
            object loDatos;
            try
            {
                Cls_Negocio_dias_trabajo obj_negocio = new Cls_Negocio_dias_trabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_Listado_dias_trabajo();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                throw ex;           
            }                     
        }
                       
        [HttpPost]
        public string ListandoOperario()
        {
            object loDatos;
            try
            {
                Cls_Negocio_dias_trabajo obj_negocio = new Cls_Negocio_dias_trabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_Listado_operario();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
                throw ex;
            }
        }


        // guardando la Data 

        [HttpPost]
        public string GuardandoInformacion( int id_dia, int  id_Operario , string NombreDia, string  HoraEntrada, string HoraSalida, int  estado, int  id_Local,  string Tipo_Mant)
        {

            bool loDatos;
            try
            {
                Cls_Entidad_dias_trabajo obj_entidad = new Cls_Entidad_dias_trabajo();
                Cls_Negocio_dias_trabajo obj_negocio = new Cls_Negocio_dias_trabajo();

                var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;
  
                obj_entidad.id_DiaTrabajo = id_dia;
                obj_entidad.id_Operario = id_Operario;
                obj_entidad.NombreDia = NombreDia;
                obj_entidad.HoraEntrada = HoraEntrada;
                obj_entidad.HoraSalida = HoraSalida;
                obj_entidad.estado = estado;
                obj_entidad.id_Local = id_Local;


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
        public string AnulandoRegistro(int id_dia, int estado_anul)
        {

            bool loDatos;
            try
            {

                Cls_Entidad_dias_trabajo obj_entidad = new Cls_Entidad_dias_trabajo();
                Cls_Negocio_dias_trabajo obj_negocio = new Cls_Negocio_dias_trabajo();

                var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

                obj_entidad.id_DiaTrabajo = id_dia;
                obj_entidad.estado = estado_anul;
                obj_entidad.usuario_edicion  = usuario;

                loDatos = obj_negocio.Capa_Negocio_Anular_Informacion(obj_entidad);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }

        // AUDITORIA


        public string AuditoriaRegistro(int id_dia)
        {

            object loDatos;
            try
            {

                Cls_Entidad_dias_trabajo obj_entidad = new Cls_Entidad_dias_trabajo();
                Cls_Negocio_dias_trabajo obj_negocio = new Cls_Negocio_dias_trabajo();

                loDatos = obj_negocio.Capa_Negocio_Get_Listado_Auditoria(id_dia);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


    }
}
