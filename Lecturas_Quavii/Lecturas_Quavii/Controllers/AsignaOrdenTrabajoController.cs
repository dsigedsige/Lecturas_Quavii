using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE.Negocio;
using DSIGE.Modelo;

namespace DSIGE.Web.Controllers
{
    public class AsignaOrdenTrabajoController : Controller
    {
        //
        // GET: /AsignaOrdenTrabajo/

        public ActionResult AsignaOrdenTrabajo()
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
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaLocales();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ListandoServicios()
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaServicioXusuario(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id); 
               // loDatos = obj_negocio.Capa_Negocio_Get_ListaServicioXusuario();
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
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaEstados();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

        [HttpPost]
        public string ListandoTecnicos(int id_local, int id_tipo_servicio)
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaTecnicos(id_local, id_tipo_servicio, ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }


        [HttpPost]
        public string ListandoOrdenesTrabajoGeneral(int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion, string horainicial, string horafinal)
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_Get_ListandoOrdenesTrabajoGeneral(((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,id_local, id_tipo_servicio, estado, suministro, medidor, tecnico, fechaAsignacion, horainicial , horafinal );
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

       [HttpPost]
        public string LecturasRelectura_Historico(string suministro, int tiposervicio)
        {
            object loDatos;
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                loDatos = obj_negocio.Capa_Negocio_LecturasRelectura_Historico(((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, suministro, tiposervicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

       [HttpPost]
       public string CortesReconexiones_Historico(string suministro, int tiposervicio)
       {
           object loDatos;
           try
           {
               Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
               loDatos = obj_negocio.Capa_Negocio_CortesReconexiones_Historico(((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, suministro, tiposervicio);
               return _Serialize(loDatos, true);
           }
           catch (Exception ex)
           {
               return _Serialize(ex.Message, true);
           }
       }




       [HttpPost]
       public string UpdateFechaAplicativoMovil(string fechamovil, int tipoServicio, List<int> List_codigos)
       {
           object loDatos = null;
           try
           {             
               var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;
               Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
               loDatos = obj_negocio.Capa_Negocio_UpdateFechaAplicativoMovil(((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, fechamovil, usuario, tipoServicio, List_codigos);
               return _Serialize(loDatos, true);

           }
           catch (Exception ex)
           {
               return _Serialize(ex.Message, true);
           }
       }

       [HttpPost]
       public string UpdateReasignaOperador(int TipoServicio, int tecnico, string fechaReasigna, List<int> List_codigos, int opcion)
       {  

           object loDatos = null;
           try
           {
               var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;
               Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
               loDatos = obj_negocio.Capa_Negocio_UpdateReasignaOperador(((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, usuario, TipoServicio, tecnico, fechaReasigna, List_codigos, opcion);
               return _Serialize(loDatos, true);

           }
           catch (Exception ex)
           {
               return _Serialize(ex.Message, true);
           }
       }
 




    }
}
