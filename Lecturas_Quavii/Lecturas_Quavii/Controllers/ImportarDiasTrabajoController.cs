using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Data.OleDb;
using System.Data;



namespace DSIGE.Web.Controllers
{
    public class ImportarDiasTrabajoController : Controller
    {
        //
        // GET: /ImportarDiasTrabajo/

        public ActionResult Inicio()
        {
            return View();
        }

        [HttpPost]
        public ActionResult JsonListarLocales()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NimportarDiasTrabajo().getLocales()),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult JsonListarSector_DiasTrabajo(int idLocal, string fechaIni, string fechaFin)
        {
            return new ContentResult
            {


                Content = MvcApplication._Serialize(new NimportarDiasTrabajo().getSector_DiasTrabajo(idLocal,fechaIni,fechaFin)),
                ContentType = "application/json"
            };
        }


        [HttpGet]
        public ActionResult JsonActualizarSector_DiasTrabajo(string sector, string fecha, string estado, int idSector)
        {
            try
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NimportarDiasTrabajo().ActualizarSector_DiasTrabajo(
                        new ImportarDiasTrabajo()
                        {                          
                            Sector = sector,
                            fecha = fecha,
                            estado = estado,
                            usuario_edicion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,                           
                            id_Sector = idSector,
                        }
                    )),
                    ContentType = "application/json"
                };
            }
            catch (Exception)
            {

                throw;
            }

        }


        [HttpPost]
        public ActionResult JsonRegistrarSector_DiasTrabajo(HttpPostedFileBase file, int idlocal)
        {
            try
            {
                           
                string extension = System.IO.Path.GetExtension(file.FileName);
                string nomExcel=  ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id + extension;
                string fileLocation = Server.MapPath("~/Upload") + "\\" + nomExcel;
         
                file.SaveAs(fileLocation);
              
              
                DataTable dt = new DataTable();
                List<ImportarDiasTrabajo> listaImportarDiasTrabajo = new List<ImportarDiasTrabajo>();
                ImportarDiasTrabajo objImportarDiasTrabajo;

                dt = new NimportarDiasTrabajo().ListaExcel(fileLocation);

                foreach (DataRow row in dt.Rows)
                {
                    objImportarDiasTrabajo = new ImportarDiasTrabajo();
                    objImportarDiasTrabajo.Sector = row["Sector"].ToString();
                    objImportarDiasTrabajo.fecha = String.Format("{0:MM/dd/yyyy}", row["fecha"].ToString());
                    objImportarDiasTrabajo.usuario_creacion = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;
                    objImportarDiasTrabajo.id_Local = idlocal;
                    listaImportarDiasTrabajo.Add(objImportarDiasTrabajo);
                }
                return new ContentResult
                    {
                    Content = MvcApplication._Serialize(new NimportarDiasTrabajo().RegistrarSector_DiasTrabajo(listaImportarDiasTrabajo)),
                    ContentType = "application/json"
                };
            }
            catch (Exception e)
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(e.Message),
                    ContentType = "application/json"
                };
            }

        }

        [HttpPost]
        public ActionResult JsonListarTablaTemp()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NLectura().NListarTablaTem(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id)),
                ContentType = "application/json"
            };
        }


    }
}
