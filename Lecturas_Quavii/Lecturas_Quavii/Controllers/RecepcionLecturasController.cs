using DSIGE.Modelo;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class RecepcionLecturasController : Controller
    {
        //
        // GET: /RecepcionLecturas/
        public ActionResult IndexProgramacionTrabajos()
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
        public string ListandoServicios()
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaServicios();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListandoOperarios()
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaOperarios();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListandoResumenTrabajos(string fecha, int id_servicio)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaTrabajosResumen(fecha, id_servicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListandoDetalleTrabajos(string fecha, int id_servicio)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaTrabajosDetalle(fecha, id_servicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



        [HttpPost]
        public string ListadoLecturasDia(int id_local, string fecha, int id_servicio, int tipoProceso, int pendiente,int asignado,int verificado)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListadoLecturasDia(id_local, fecha, id_servicio, tipoProceso, pendiente,asignado, verificado);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListadoOperariosZona(int id_local, string fecha, int id_servicio, int tipoProceso)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListadoOperariosZona(id_local, fecha, id_servicio, tipoProceso);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string GenerarSecuencia(int idLecturista,int idServicios,string idSector, string FechaInicial,int Ano,int Mes)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocios_GenerarSecuencia(idLecturista, idServicios, idSector, FechaInicial,Ano,Mes);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        
        [HttpPost]
        public string GuardarLecturasDia(List<Cls_Entidad_Temporal_Sum> listLecturas, int id_operario, int tipoServicio,string fecha_asignacion)
        {
            object loDatos;
            try
            {
                foreach (var item in listLecturas)
                {
                    item.id_Operario = id_operario;
                    item.Usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;
                }
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocios_GuardarTemporalSum(listLecturas,tipoServicio,fecha_asignacion);                
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



        [HttpPost]
        public string ListadoResumenOperario(string fecha, int id_Operario)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListadoResumenOperario(fecha, id_Operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ListandoEstados()
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaEstados();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }
        
        [HttpPost]
        public string ListandoInformacionLecturas(int local, string fechaAsigna, int servicio, int id_operador, string rt, string s1, string sector, string zona,string suministro, string medidor, int estado)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaInformacionLecturas(local, fechaAsigna, servicio, id_operador, rt, s1, sector, zona, suministro, medidor, estado);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }




        [HttpPost]
        public string ListandoInformacionLecturas_new(int local, string fechaAsigna, int servicio, int id_operador, string rt, string s1, string sector, string zona, string suministro, string medidor, int estado, int id_zona)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaInformacionLecturas_new(local, fechaAsigna, servicio, id_operador, rt, s1, sector, zona, suministro, medidor, estado, id_zona);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        
        [HttpPost]
        public string Lista_Lecturas_Bak(string fecha_movil,string fecha_servidor)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaInformacionLecturas_Bak(fecha_movil, fecha_servidor);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }




        [HttpPost]
        public string AnularLectura(int id_lectura)
        {
            object loDatos;
            try
            {
               
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocios_AnularLectura(id_lectura);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }




        public string Update_Lecturas(int id_lectura, string direccion, string lectura_actual, string observacion, string medidor, string dns, string ubicacion)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_Update_Lectura(id_lectura, direccion, lectura_actual, observacion, medidor, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, dns, ubicacion);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        public string Update_Operarios_Historial(string listIds)
        {
            object loDatos;
            try
            {
                RecepcionLecturas_BL obj_negocio = new RecepcionLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Update_operarios_Historial(listIds);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


    }
}
