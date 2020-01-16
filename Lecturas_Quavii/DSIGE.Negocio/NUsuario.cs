using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Dato;
using System.Net.Mail;
using System.Net;

namespace DSIGE.Negocio
{
    public class NUsuario
    {

        public NUsuario()
        {
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 27-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Usuario NRecuperaClave(Request_Usuario_Recupera_Clave oRq)
        {
            try
            {
                return new DUsuario().DRecuperaClave(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Usuario> getUsuarios()
        {
            try
            {
                return new DUsuario().getUsuarios();
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public List<Cargo_Personal> getCargo()
        {
            try
            {
                return new DUsuario().getCargo();
            }
            catch (Exception e)
            {

                throw e;
            }
        }


        public List<Perfil> getPerfil()
        {
            try
            {
                return new DUsuario().gePerfil();
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public bool NregistrarUsu(Usuario objUsu)
        {
            try
            {
                return new DUsuario().InsertarUsuario(objUsu);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool NActualizarUsu(Usuario objUsu)
        {
            try
            {
                return new DUsuario().ActualizarUsuario(objUsu);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool NAnularUsu(Usuario objUsu)
        {
            try
            {
                return new DUsuario().AnularUsuario(objUsu);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Usuario> ListarUsuarios()
        {
            try
            {
                return new DUsuario().ListarUsuarios();
            }
            catch (Exception e)
            {

                throw e;
            }

        }

        public List<Usuario> GetDatosUsuario(int idUsuario)
        {
            return new DUsuario().GetDatosUsuario(idUsuario);
        }


        public bool CmbiarPassUsuario(Usuario objUsuario)
        {
            return new DUsuario().CmbiarPassUsuario(objUsuario);
        }

        


        //public int NEnvioEmail(string vHost, int vPuerto, string vUsuario, string vContraseña, string[] MensajePrueba)
        //{
        //    Usuario usuario = DUsuario().DRecuperaClave();

        //    try
        //    {
        //        var body = "<p>Email From: {0} ({1})</p><p>Message:</p><p>{2}</p>";

        //        MailMessage oMail = new MailMessage();
        //        SmtpClient oSmtp = new SmtpClient();

        //        oSmtp.Host = vHost;
        //        oSmtp.Port = vPuerto;
        //        oSmtp.Credentials = new NetworkCredential(vUsuario, vContraseña);
        //        oSmtp.EnableSsl = true;
        //        oMail.From = new MailAddress(vUsuario);

        //        Usuario usu = new Usuario();

        //        oMail.To.Add(usu.usu_email);


        //        //oMail.Subject = "Relevo de Precios " + oPeriodo.rpl_descripcion;
        //        oMail.Subject = "INFORME DE PRECIO - Relevo";
        //        oMail.IsBodyHtml = true;
        //        oMail.Body = body;

        //        oSmtp.Send(oMail);

        //        oSmtp = null;
        //        oMail = null;

        //        return 1;
        //    }
        //    catch (Exception)
        //    {
        //        return 0;
        //    }

        //}
        public bool Capa_Negocio_ModificarCrear_UsuarioServicio(int estado, List<int> List_codigos)
        {
            try
            {
                DUsuario Objeto_Dato = new DUsuario();



                Cls_Dato_AsignaOrdenTrabajo Objeto_Servicio = new Cls_Dato_AsignaOrdenTrabajo();
                var serv = Objeto_Servicio.Capa_Dato_Get_ListaServicio();


                foreach (int elemet in List_codigos)
                {
                    foreach (var item01 in serv)
                    {

                        Objeto_Dato.Capa_Dato_Modificar_UsuarioServicio(elemet, Convert.ToInt32(item01.id_TipoServicio), estado);
                        
                    }



                }


                return true;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }





        public bool Capa_Negocio_Modificar_UsuarioServicio(int idusuario, int idtiposervicio, int estadousuario)
        {
            try
            {
                DUsuario Objeto_Dato = new DUsuario();
                return Objeto_Dato.Capa_Dato_Modificar_UsuarioServicio(idusuario, idtiposervicio, estadousuario);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public List<Usuario_Servicio> Capa_Negocio_Listar_UsuarioxServ(int idusuario)
        {
            try
            {
                DUsuario Objeto_Dato = new DUsuario();
                return Objeto_Dato.Capa_Dato_ListarUsuarioServ(idusuario);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }



        

    }
}
