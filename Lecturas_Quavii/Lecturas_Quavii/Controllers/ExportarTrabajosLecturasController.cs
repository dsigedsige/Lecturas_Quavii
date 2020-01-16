
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System.IO;

using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Configuration;
using System.Drawing;
using System.Globalization;
using System.Data;

namespace DSIGE.Web.Controllers
{
    public class ExportarTrabajosLecturasController : Controller
    {
        //
        // GET: /ExportarTrabajosLecturas/

        public ActionResult ExportarTrabajosLecturas()
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
      public string MostrarInformacion(string fechaAsignacion, int TipoServicio )
        {
            object loDatos;
            try
            {
                Cls_Negocio_Export_trabajos_lectura obj_negocio = new Cls_Negocio_Export_trabajos_lectura();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaLecturas(fechaAsignacion, TipoServicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

      [HttpPost]
      public string DescargaExcel(string fechaAsignacion, int TipoServicio)
      {
            int _fila = 2;
            string _ruta;
            string nombreArchivo = "";
            string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];
            var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

          try
          {           
              List<Cls_Entidad_Export_trabajos_lectura> _lista = new List<Cls_Entidad_Export_trabajos_lectura>();

              Cls_Negocio_Export_trabajos_lectura obj_negocio = new DSIGE.Negocio.Cls_Negocio_Export_trabajos_lectura();
              _lista = obj_negocio.Capa_Negocio_Get_ListaLecturas_Excel(fechaAsignacion, TipoServicio);
              
              if (_lista.Count == 0)
              {
                  return _Serialize("0|No hay informacion para mostrar.", true);
              }

                if (TipoServicio==1)
                {
                    nombreArchivo = "LECTURAS_EXPORTADO" + usuario + ".xls";
                }
                else if (TipoServicio == 2)
                {
                    nombreArchivo = "RELECTURAS_EXPORTADO" + usuario + ".xls";
                }
                else if (TipoServicio == 9)
                {
                    nombreArchivo = "RECLAMOS_EXPORTADO_" + usuario + ".xls";
                }

                _ruta = Path.Combine(Server.MapPath("~/Temp") + "\\" + nombreArchivo );

              FileInfo _file = new FileInfo(_ruta);
              if (_file.Exists)
              {
                  _file.Delete();
                  _file = new FileInfo(_ruta);
              }

              using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
              {
                  Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Importar");
                    oWs.Cells.Style.Font.SetFromFont(new Font("Tahoma", 8));
                    for (int i = 1; i <= 21; i++)
                    {
                        oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        oWs.Cells[1, i].Style.Font.Size = 9; //letra tamaño   
                        oWs.Cells[1, i].Style.Font.Bold = true; //Letra negrita 
                    } 

                    oWs.Cells[1, 1].Value = "ITEM";
                    oWs.Cells[1, 2].Value = "NOBORAR";
                    oWs.Cells[1, 3].Value = "INSTALACIÓN";                    
                    oWs.Cells[1, 4].Value = "APARATO";

                    oWs.Cells[1, 5].Value = "TIPO CALLE";
                    oWs.Cells[1, 6].Value = "NOMBRE DE CALLE";
                    oWs.Cells[1, 7].Value = "ALTURA DE CALLE";
                    oWs.Cells[1, 8].Value = "NÚMERO DE EDIFICIO";
                    oWs.Cells[1, 9].Value = "NÚMERO DE DEPARTAMENTO";
 
                    oWs.Cells[1, 10].Value = "DETALLE CONSTRUCCIÓN (OBJETO DE CONEXIÓN)";
                    oWs.Cells[1, 11].Value = "CONJUNTO DE VIVIENDA (OBJETO DE CONEXIÓN)";
                    oWs.Cells[1, 12].Value = "MANZANA/LOTE";
                    oWs.Cells[1, 13].Value = "DISTRITO";

                    oWs.Cells[1, 14].Value = "CUENTA CONTRATO";
                    oWs.Cells[1, 15].Value = "SECUENCIA DE LECTURA";
                    oWs.Cells[1, 16].Value = "UNIDAD DE LECTURA";
                    oWs.Cells[1, 17].Value = "NÚMERO DE LECTURAS ESTIMADAS CONSECUTIVAS";
 
                    oWs.Cells[1, 18].Value = "EMPRESA LECTORA"; 
                    oWs.Cells[1, 19].Value = "NOTA 2 DE LA UBICACIÓN DEL APARATO";
                    oWs.Cells[1, 20].Value = "TECNICO";
                    oWs.Cells[1, 21].Value = "SECUENCIA"; 

                    int acu = 0;
                   foreach (Cls_Entidad_Export_trabajos_lectura oBj in _lista)
                    {
                        acu = acu + 1 ;

                        for (int i = 1; i <= 21; i++)
                        {
                            oWs.Cells[_fila, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }

                        oWs.Cells[_fila,1].Value = acu;
                        oWs.Cells[_fila,2].Value = oBj.id_Lectura;
                        oWs.Cells[_fila,3].Value= oBj.Instalacion;

                        //oWs.Cells[_fila, 4].Style.Numberformat.Format = "#,##0";
                        //oWs.Cells[_fila, 4].Value = Convert.ToDouble(oBj.Aparato);
                        oWs.Cells[_fila, 4].Value =  oBj.Aparato;

                        oWs.Cells[_fila, 5].Value = oBj.Tipo_calle;
                        oWs.Cells[_fila, 6].Value = oBj.Nombre_Calle;
                        oWs.Cells[_fila, 7].Value = oBj.Altura_Calle;
                        oWs.Cells[_fila, 8].Value = oBj.Numero_Edificio;
                        oWs.Cells[_fila, 9].Value = oBj.Numero_Departamento; 
 
                        oWs.Cells[_fila, 10].Value = oBj.Detalle_Construccion;
                        oWs.Cells[_fila, 11].Value = oBj.Conjunto_Vivienda;
                        oWs.Cells[_fila, 12].Value = oBj.Manzana_Lote;
                        oWs.Cells[_fila, 13].Value = oBj.Distrito;

                        oWs.Cells[_fila, 14].Value = oBj.Cuenta_contrato;
                        oWs.Cells[_fila, 15].Value = oBj.Secuencia_lectura;
                        oWs.Cells[_fila, 16].Value = oBj.Unidad_lectura;
                        oWs.Cells[_fila, 17].Value = oBj.Numero_lecturas_estimadas_consecutivas; 
                        oWs.Cells[_fila, 18].Value = oBj.Empresa_Lectora ;
 
                        oWs.Cells[_fila, 19].Value = oBj.Nota_2_ubicacion_aparato;
                        oWs.Cells[_fila, 20].Value = oBj.Tecnico ;
                        oWs.Cells[_fila, 21].Value = oBj.Secuencia; 

                      _fila++;
                  }

                oWs.Row(1).Style.Font.Bold = true;
                oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;

                oWs.Column(1).Style.Font.Bold = true;

                for (int i = 1; i <=21; i++)
                {
                    oWs.Column(i).AutoFit(); 
                }
            
                  oEx.Save();
              }


              return _Serialize("1|" + ruta_descarga+ nombreArchivo, true);


          }
          catch (Exception ex )
          {
          
            return _Serialize("0|" + ex.Message, true);
          }           

      }


        [HttpPost]
        public string DescargaExcel_New(int Local, string fechaAsignacion, int TipoServicio)
        {
            int _fila = 5;
            string _ruta;
            string nombreArchivo = "";
            string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];
            var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

            try
            {


                DataTable dt_detalles = new DataTable();

                Cls_Negocio_Export_trabajos_lectura obj_negocio = new DSIGE.Negocio.Cls_Negocio_Export_trabajos_lectura();
                dt_detalles = obj_negocio.Capa_Negocio_Get_ListaCortesReconexion_Excel(Local,fechaAsignacion, TipoServicio);

                if (dt_detalles.Rows.Count <= 0)
                {
                    return _Serialize("0|No hay informacion para mostrar.", true);
                }

                if (TipoServicio == 3)
                {
                    nombreArchivo = "CORTES_EXPORTADO" + usuario + ".xls";
                }
                else if (TipoServicio == 4)
                {
                    nombreArchivo = "RECONEXION_EXPORTADO" + usuario + ".xls";
                }
 
                _ruta = Path.Combine(Server.MapPath("~/Temp") + "\\" + nombreArchivo);

                FileInfo _file = new FileInfo(_ruta);
                if (_file.Exists)
                {
                    _file.Delete();
                    _file = new FileInfo(_ruta);
                }

                using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
                {
                    Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Estructura");
                    oWs.Cells.Style.Font.SetFromFont(new Font("Tahoma", 8));
                                                                          
                    oWs.Cells[1, 1].Value = "ESTRUCTURA LEGALIZACIÓN MASIVA DE ORDENES DE CORTES";
                    oWs.Cells[1, 1].Style.Font.Size = 18; //letra tamaño   
                    oWs.Cells[1, 1, 1, 31].Merge = true;  // combinar celdaS dt
                    oWs.Cells[1, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[1, 1].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[1, 1].Style.Font.Bold = true; //Letra negrita

                    oWs.Cells[3, 5].Value = "Actividades";
                    oWs.Cells[3, 5].Style.Font.Size = 10; //letra tamaño   
                    oWs.Cells[3, 5, 3, 19].Merge = true;  // combinar celdaS dt
                    oWs.Cells[3, 5].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[3, 5].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[3, 5].Style.Font.Bold = true; //Letra negrita
                    
                    oWs.Cells[3, 21].Value = "Lecturas_Elementos";
                    oWs.Cells[3, 21].Style.Font.Size = 10; //letra tamaño   
                    oWs.Cells[3, 21, 3, 27].Merge = true;  // combinar celdaS dt
                    oWs.Cells[3, 21].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[3, 21].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[3, 21].Style.Font.Bold = true; //Letra negrita
                     
                    oWs.Cells[3, 28].Value = "Tipo Comentario";
                    oWs.Cells[3, 28].Style.Font.Size = 10; //letra tamaño   
                    oWs.Cells[3, 28, 3, 29].Merge = true;  // combinar celdaS dt
                    oWs.Cells[3, 28].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[3, 28].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[3, 28].Style.Font.Bold = true; //Letra negrita
                                       
                    oWs.Cells[3, 30].Value = "Fechas Legalización";
                    oWs.Cells[3, 30].Style.Font.Size = 10; //letra tamaño   
                    oWs.Cells[3, 30, 3, 31].Merge = true;  // combinar celdaS dt
                    oWs.Cells[3, 30].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[3, 30].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[3, 30].Style.Font.Bold = true; //Letra negrita



                    oWs.Cells[4, 1].Value = "Orden";
                    oWs.Cells[4, 2].Value = "Causal";
                    oWs.Cells[4, 3].Value = "Personal";
                    oWs.Cells[4, 4].Value = "Datos_Adicionales";
                    oWs.Cells[4, 5].Value = "Or_Activity_id>";
                    oWs.Cells[4, 6].Value = "Cant Legal;";
                    oWs.Cells[4, 7].Value = "Nombre Atributo 1>";
                    oWs.Cells[4, 8].Value = "Medidor>";
                    oWs.Cells[4, 9].Value = "Id Componente>";
                    oWs.Cells[4, 10].Value = "Sello1=";

                    oWs.Cells[4, 11].Value = "Codigo Ubicación Sello=";
                    oWs.Cells[4, 12].Value = "Acción=";
                    oWs.Cells[4, 13].Value = "Manipulado S/N=";
                    oWs.Cells[4, 14].Value = "Medidor!";
                    oWs.Cells[4, 15].Value = "Sello2=";

                    oWs.Cells[4, 16].Value = "Ubicación=";
                    oWs.Cells[4, 17].Value = "Acción=";
                    oWs.Cells[4, 18].Value = "Manipulado S/N=";
                    oWs.Cells[4, 19].Value = "Medidor;;;";
                    oWs.Cells[4, 20].Value = "Items_Elementos";

                    oWs.Cells[4, 21].Value = "Medidor;";
                    oWs.Cells[4, 22].Value = "Consumo=";
                    oWs.Cells[4, 23].Value = "Lectura=";
                    oWs.Cells[4, 24].Value = "Causa=";
                    oWs.Cells[4, 25].Value = "Observación1=";

                    oWs.Cells[4, 26].Value = "Observación2=";
                    oWs.Cells[4, 27].Value = "Observación3";
                    oWs.Cells[4, 28].Value = "Código Comentario;";
                    oWs.Cells[4, 29].Value = "Comentario";
                    oWs.Cells[4, 30].Value = "Fecha Ini Ejec;";

                    oWs.Cells[4, 31].Value = "Fecha Fin Ejec";


                    int acu = 0;
                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        acu = acu + 1;
                        oWs.Cells[_fila, 1].Value =  oBj["Orden"].ToString();
                        oWs.Cells[_fila, 2].Value =  oBj["Causal"].ToString();
                        oWs.Cells[_fila, 3].Value =  oBj["Personal"].ToString();
                        oWs.Cells[_fila, 4].Value =  oBj["Datos_Adicionales"].ToString();
                        oWs.Cells[_fila, 5].Value =  oBj["Or_Activity_id"].ToString();
                        oWs.Cells[_fila, 6].Value =  oBj["CantLegal"].ToString();
                        oWs.Cells[_fila, 7].Value =  oBj["NombreAtributo1"].ToString();
                        oWs.Cells[_fila, 8].Value =  oBj["Medidor1"].ToString();
                        oWs.Cells[_fila, 9].Value =  oBj["IdComponente"].ToString();
                        oWs.Cells[_fila, 10].Value = oBj["Sello1"].ToString();

                        oWs.Cells[_fila, 11].Value = oBj["CodigoUbicacionSello"].ToString();
                        oWs.Cells[_fila, 12].Value = oBj["Accion1"].ToString();
                        oWs.Cells[_fila, 13].Value = oBj["ManipuladoS_N1"].ToString();
                        oWs.Cells[_fila, 14].Value = oBj["Medidor2"].ToString();
                        oWs.Cells[_fila, 15].Value = oBj["Sello2"].ToString();

                        oWs.Cells[_fila, 16].Value = oBj["Ubicacion"].ToString();
                        oWs.Cells[_fila, 17].Value = oBj["Accion2"].ToString();
                        oWs.Cells[_fila, 18].Value = oBj["ManipuladoS_N2"].ToString();
                        oWs.Cells[_fila, 19].Value = oBj["Medidor3"].ToString();
                        oWs.Cells[_fila, 20].Value = oBj["Items_Elementos"].ToString();

                        oWs.Cells[_fila, 21].Value = oBj["Medidor4"].ToString();
                        oWs.Cells[_fila, 22].Value = oBj["Consumo"].ToString();
                        oWs.Cells[_fila, 23].Value = oBj["Lectura"].ToString();
                        oWs.Cells[_fila, 24].Value = oBj["Causa"].ToString();
                        oWs.Cells[_fila, 25].Value = oBj["Observacion1"].ToString();

                        oWs.Cells[_fila, 26].Value = oBj["Observacion2"].ToString();
                        oWs.Cells[_fila, 27].Value = oBj["Observacion3"].ToString();
                        oWs.Cells[_fila, 28].Value = oBj["CodigoComentario"].ToString();
                        oWs.Cells[_fila, 29].Value = oBj["Comentario"].ToString();
                        oWs.Cells[_fila, 30].Value = oBj["FechaIniEjec"].ToString();
                        oWs.Cells[_fila, 31].Value = oBj["FechaFinEjec"].ToString();

                        _fila++;
                    }

                    oWs.Row(4).Style.Font.Bold = true;
                    oWs.Row(4).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(4).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;           

                    for (int i = 1; i <= 31; i++)
                    {
                        oWs.Column(i).AutoFit();
                    }

                    oEx.Save();
                }
                return _Serialize("1|" + ruta_descarga + nombreArchivo, true);
            }
            catch (Exception ex)
            {
                return _Serialize("0|" + ex.Message, true);
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
