using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
    public class NEnvioTrabajoCliLecReLec
    {
        public NEnvioTrabajoCliLecReLec() { }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 20-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        /// 



        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Negocio_Excel_Envio_Trabajo_cliente(int id_Local ,int id_operario, string suministro_lectura, string medidor_lectura, string fechaAsignacion_Lectura, int estado)
        {

            try
            {
                DEnvioTrabajoCliLecReLec Objeto_Dato = new DEnvioTrabajoCliLecReLec();
                return Objeto_Dato.Capa_Dato_Excel_Envio_Trabajo_cliente(id_Local, id_operario,suministro_lectura, medidor_lectura, fechaAsignacion_Lectura, estado);
            }
            catch (Exception)
            {

                throw;
            }


        }

        public bool Capa_Negocio_Guardar_Informacion(int id_Lectura, string id_Operario_Lectura, string Lectura_Manual_Lectura, DateTime fechaLecturaMovil_lectura, string id_Observacion_Lectura, int id_usuario)
        {
            try
            {
                DEnvioTrabajoCliLecReLec Objeto_Dato = new DEnvioTrabajoCliLecReLec();
                return Objeto_Dato.Capa_Dato_Guardar_Informacion(id_Lectura, id_Operario_Lectura, Lectura_Manual_Lectura,fechaLecturaMovil_lectura, id_Observacion_Lectura, id_usuario);
            }
            catch (Exception)
            {

                throw;
            }

        }



        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Negocio_Listado_Operario()
        {

            try
            {
                DEnvioTrabajoCliLecReLec Objeto_Dato = new DEnvioTrabajoCliLecReLec();
                return Objeto_Dato.Capa_Dato_Listado_Operario();
            }
            catch (Exception)
            {

                throw ;
            }


        }


        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Negocio_Listado_Observaciones()
        {

            try
            {
                DEnvioTrabajoCliLecReLec Objeto_Dato = new DEnvioTrabajoCliLecReLec();
                return Objeto_Dato.Capa_Dato_Listado_Observaciones();
            }
            catch (Exception)
            {

                throw;
            }


        }



        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Negocio_Listando_Registros_Modo_edicion(int id_lectura)
        {

            try
            {
                DEnvioTrabajoCliLecReLec Objeto_Dato = new DEnvioTrabajoCliLecReLec();
                return Objeto_Dato.Capa_Dato_Listando_Registros_Modo_edicion(id_lectura);
            }
            catch (Exception)
            {

                throw;
            }


        }













        public List<EnvioTrabajoCliLecReLec> NLista(Request_Asigna_Lectura_ReLectura oRq)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DLista(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 18-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NActualizaEnviaCliente(string idLectura, int usuEdi, int idEmp)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DActualizaEnviaCliente(idLectura, usuEdi, idEmp);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 28-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<EnvioTrabajoCliLecReLec> NListaFotos(int idLectura)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DListaFotos(idLectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Fotoselfie_reparto> NFoto_selfie_reparto(string fecha, int id_operario)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DFotosSelfie_Reparto(fecha, id_operario );
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<EnvioTrabajoCliLecReLec> NListaFotos_Historico(int idLectura , int Tiposervicio)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DListaFotos_Historico(idLectura, Tiposervicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<EnvioTrabajoCliLecReLec> NListaPreVisualizaEnvio(string codLectura)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DListaPreVisualizaEnvio(codLectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<EnvioTrabajoCliLecReLec> NListaPreVisualizLecturaEnvioCliente(string fechamovil, int tipoServicio, List<int> List_codigos)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DListaPreVisualizaLecturaEnvioCliente(fechamovil, tipoServicio, List_codigos);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<EnvioTrabajoCliLecReLecLecturaPendiente> NListaLecturasPendientes(DateTime fechaAsigna, int idEmp)
        {
            try
            {
                return new DEnvioTrabajoCliLecReLec().DListaLecturasPendientes(fechaAsigna, idEmp);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
