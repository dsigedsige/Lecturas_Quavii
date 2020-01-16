using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Data.OleDb;
using System.Data;
using Newtonsoft.Json;
using System.Text;

namespace DSIGE.Web.Controllers
{
    public class Importar_ArchivoTXTController : Controller
    {

        public ActionResult Inicio()
        {
            return View();
        }

        public ActionResult inicio_index()
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
        public ActionResult saveTextToServidor(HttpPostedFileBase file, string fechaAsignacion, string TipoServicio)
        {

            List<CorteTemporalCorte> oCortes = new List<CorteTemporalCorte>();
            DateTime _fecha_actual = DateTime.Now;
             
                object loDatos = null;
                string nomFile = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id + file.FileName;
                string fileLocation = Server.MapPath("~/Upload") + "\\" + nomFile;
                file.SaveAs(fileLocation);

                List<ImportarArchivoPlano> List_obj_datos = new List<ImportarArchivoPlano>();
                 string[] lines = System.IO.File.ReadAllLines(fileLocation, Encoding.Default);
                 foreach (string line in lines)
                 {
                     var arrayText = line.ToString().Split('\t'); 

                     ImportarArchivoPlano obj_entidad = new ImportarArchivoPlano();
                        obj_entidad.Item= arrayText[0];
                        obj_entidad.Instalacion= arrayText[1];
                        obj_entidad.Equipo= arrayText[2];
                        obj_entidad.Aparato= arrayText[3];
                        obj_entidad.Tipo_calle= arrayText[4];
                        obj_entidad.Nombre_Calle= arrayText[5];
                        obj_entidad.Altura_Calle= arrayText[6];
                        obj_entidad.Numero_Edificio= arrayText[7];
                        obj_entidad.Numero_Departamento= arrayText[8];
                        obj_entidad.Detalle_adicional_ubicacion= arrayText[9];
                        obj_entidad.Piso= arrayText[10];
                        obj_entidad.Vivienda_Principal= arrayText[11];
                        obj_entidad.Detalle_Construccion= arrayText[12];
                        obj_entidad.Conjunto_Vivienda= arrayText[13];
                        obj_entidad.Manzana_Lote= arrayText[14];
                        obj_entidad.Distrito= arrayText[15];
                        obj_entidad.Codigo_postal= arrayText[16];
                        obj_entidad.Poblacion= arrayText[17];
                        obj_entidad.Emplazamiento= arrayText[18];
                        obj_entidad.Suplemento_emplazamiento= arrayText[19];
                        obj_entidad.Lectura_anterior= "0";
                        obj_entidad.Fecha_planificada_lectura_anterior= arrayText[20];
                        obj_entidad.Fecha_planificada_lectura_actual= arrayText[21];
                        obj_entidad.Fecha_planificada_lectura_proxima= arrayText[22];
                        obj_entidad.Interlocutor_comercial= arrayText[23];
                        obj_entidad.Cuenta_contrato= arrayText[24];
                        obj_entidad.Tipo_Cliente= arrayText[25];
                        obj_entidad.Categoria= arrayText[26];
                        obj_entidad.Secuencia_lectura= arrayText[27];
                        obj_entidad.Unidad_lectura= arrayText[28];
                        obj_entidad.Numero_lecturas_estimadas= arrayText[29];
                        obj_entidad.Marca_primera_lectura= arrayText[30];
                        obj_entidad.Empresa_Lectora= arrayText[31];
                        obj_entidad.Campo1= "";
                        obj_entidad.Campo2= "";
                        obj_entidad.Cliente= arrayText[34];
                        obj_entidad.Nota_ubicacion_aparato= arrayText[35];
                        obj_entidad.Nota_dos_ubicacion_aparato= arrayText[36];
                        obj_entidad.Secuencia= arrayText[37];
                        obj_entidad.nombre_ArchivoImportado = file.FileName;
                        obj_entidad.fecha_Asignacion= fechaAsignacion;
                        obj_entidad.id_TipoServicio = TipoServicio;

                     List_obj_datos.Add(obj_entidad);  
                 }

                 ImportarArchivoPlano_BL obj_negocio = new ImportarArchivoPlano_BL();
                 string obj = obj_negocio.Capa_Negocio_GuardarArchivos(List_obj_datos, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id);

                 return new ContentResult
                 {
                     Content = MvcApplication._Serialize(obj),
                     ContentType = "application/json"
                 };                         
        }


        [HttpPost]
        public ActionResult saveTextToServidor_new(HttpPostedFileBase file, string fechaAsignacion, string TipoServicio)
        {
            object loDatos = null;
 
            try
            {
                List<CorteTemporalCorte> oCortes = new List<CorteTemporalCorte>();
                DateTime _fecha_actual = DateTime.Now;

                string nomFile = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id + file.FileName;
                string fileLocation = Server.MapPath("~/Upload") + "\\" + nomFile;
                file.SaveAs(fileLocation);

                List<ImportarArchivoPlano_new> List_obj_datos = new List<ImportarArchivoPlano_new>();
                string[] lines = System.IO.File.ReadAllLines(fileLocation,   Encoding.Default);
           

                foreach (string line in lines)
                {
                    var arrayText = line.ToString().Split('\t');

                    ImportarArchivoPlano_new obj_entidad = new ImportarArchivoPlano_new();
                    obj_entidad.Item = arrayText[0];
                    obj_entidad.Instalacion = arrayText[1];
                    obj_entidad.Equipo = arrayText[2];
                    obj_entidad.Aparato = arrayText[3];
                    obj_entidad.Tipo_calle = arrayText[4];
                    obj_entidad.Nombre_Calle = arrayText[5];
                    obj_entidad.Altura_Calle = arrayText[6];
                    obj_entidad.Numero_Edificio = arrayText[7];
                    obj_entidad.Numero_Departamento = arrayText[8];
                    obj_entidad.Detalle_adicional_ubicacion = arrayText[9];
                    obj_entidad.Piso = arrayText[10];
                    obj_entidad.Vivienda_Principal = arrayText[11];
                    obj_entidad.Detalle_Construccion = arrayText[12];
                    obj_entidad.Conjunto_Vivienda = arrayText[13];
                    obj_entidad.Manzana_Lote = arrayText[14];
                    obj_entidad.Distrito = arrayText[15];
                    obj_entidad.Codigo_postal = arrayText[16];
                    obj_entidad.Poblacion = arrayText[17];
                    obj_entidad.Emplazamiento = arrayText[18];
                    obj_entidad.Suplemento_emplazamiento = arrayText[19];
                    obj_entidad.Lectura_anterior = "0";
                    obj_entidad.Fecha_planificada_lectura_anterior = arrayText[20];
                    obj_entidad.Fecha_planificada_lectura_actual = arrayText[21];
                    obj_entidad.Fecha_planificada_lectura_proxima = arrayText[22];
                    obj_entidad.Interlocutor_comercial = arrayText[23];
                    obj_entidad.Cuenta_contrato = arrayText[24];
                    obj_entidad.Tipo_Cliente = arrayText[25];
                    obj_entidad.Categoria = arrayText[26];
                    obj_entidad.Secuencia_lectura = arrayText[27];
                    obj_entidad.Unidad_lectura = arrayText[28];
                    obj_entidad.Numero_lecturas_estimadas = arrayText[29];
                    obj_entidad.Marca_primera_lectura = arrayText[30];
                    obj_entidad.Empresa_Lectora = arrayText[31];
                    obj_entidad.Campo1 = "";
                    obj_entidad.Campo2 = "";
                    obj_entidad.Cliente = arrayText[34];
                    obj_entidad.Nota_ubicacion_aparato = arrayText[35];
                    obj_entidad.Nota_dos_ubicacion_aparato = arrayText[36];
                    obj_entidad.Secuencia = arrayText[37];
                    obj_entidad.estado_instalacion = arrayText[38];
                    obj_entidad.lectura_corte = arrayText[39];
                    obj_entidad.lectura_maxima = arrayText[40];
                    obj_entidad.lectura_minima = arrayText[41];
                    obj_entidad.estado_contrato = arrayText[42];
                    obj_entidad.lectura_inmediata_ant = arrayText[43];

                    obj_entidad.latitud = arrayText[44];
                    obj_entidad.longitud = arrayText[45];

                    obj_entidad.nombre_ArchivoImportado = file.FileName;
                    obj_entidad.fecha_Asignacion = fechaAsignacion;
                    obj_entidad.id_TipoServicio = TipoServicio;

                    List_obj_datos.Add(obj_entidad);
                }

                ImportarArchivoPlano_BL obj_negocio = new ImportarArchivoPlano_BL();
                string obj = obj_negocio.Capa_Negocio_GuardarArchivos_new(List_obj_datos, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id);

                return new ContentResult
                {
                    Content = MvcApplication._Serialize(obj),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {
                throw;
            }

        }


        [HttpPost]
        public string VerificandoArchivos(string fechaAsignacion, int TipoServicio)
        {
            object loDatos;
            try
            {
                ImportarArchivoPlano_BL obj_negocio = new ImportarArchivoPlano_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaArchivosAlmacenados(fechaAsignacion, TipoServicio);
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
                loDatos = obj_negocio.Capa_Negocio_Get_ListaServicioXusuario_II(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

    }
}
