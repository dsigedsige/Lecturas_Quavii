using DSIGE.Modelo;
using DSIGE.Negocio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.IO;
namespace DSIGE.Web.Controllers
{
    public class Listado_Lecturas_FotosController : Controller
    {
        //
        // GET: /Listado_Lecturas_Fotos/

        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <returns></returns>

        [HttpPost]
        public ActionResult FotosLectura(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NEnvioTrabajoCliLecReLec().NListaFotos(
                            Convert.ToInt32(__a)
                        )
                    , true),
                ContentType = "application/json"
            };
        }
        
 

        [HttpPost]
        public ActionResult JsonListaLecturaTomadaFoto(string __a, string __c, string __d, string __e, string __f,string lista, int id_operario)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NLectura().NListaLecturaTomadaFoto(
                            new Request_Lectura_Tomadas()
                            {
                                local = Convert.ToInt32(__a),
                                //tipo = Convert.ToInt32(__b),
                                f_ini = Convert.ToString(__c),
                                f_fin = Convert.ToString(__d),
                                suministro = Convert.ToString(__e),
                                medidor = Convert.ToString(__f)
                            }, lista, id_operario
                        )
                    , true),
                ContentType = "application/json"
            };
        }

        public ActionResult JsonListaURlLectura(string id,string fechaini,string fechafin,int con)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NLectura().NListaUrlLectura(id,fechaini,fechafin,con)
                    , true),
                ContentType = "application/json"
            };
        }

        public ActionResult dowloadFoto(string urlArray ){
        
            string[] urls;
            urls = urlArray.Split(',');
            string folderName = @"C:\";
            string folderName2 = @"C:\fotos";

        // To create a string that specifies the path to a subfolder under your 
        // top-level folder, add a name for the subfolder to folderName.
            string pathString = System.IO.Path.Combine(folderName, "fotos");
            string pathString2 = System.IO.Path.Combine(folderName2, "dsigee");
            
        Directory.CreateDirectory(pathString);
        Directory.CreateDirectory(pathString2);
        System.Diagnostics.Process.Start(@"C:\fotos\dsigee");
            for (int i = 0; i < urls.Length; i++)
            {
                string[] suburls;
                suburls = urls[i].Split('|');
                string localFilename = @"C:\fotos\dsigee\" + suburls[1] + "_" + suburls[2] + ".jpg";

                using (WebClient client = new WebClient())
                {
                    client.DownloadFile("http://www.dsige.com/Calidda/Content/foto/foto/" + suburls[0], localFilename);
                    
                  
                }
        
            }
               
            
                return new ContentResult { 
                    Content=MvcApplication._Serialize("Completo", true),
                    ContentType="application/json"
                };
        }
      

    }
}
