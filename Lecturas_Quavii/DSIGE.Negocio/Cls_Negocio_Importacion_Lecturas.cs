using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class Cls_Negocio_Importacion_Lecturas
    {
        //ACTUALIZANDO FECHA DE LECTURA

        public bool Capa_Negocio_ListaExcel_Actualizar_fecha(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_ListaExcel_Actualizar_fecha(fileLocation, usuario, idlocal, idservicio, idfechaAsignacion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_Importacion_Lecturas> Capa_Negocio_Listar_Temporal_Actualizar_Fecha(int cod_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Datos_Listar_Temporal_Actualizar_Fecha(cod_usuario);
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public bool Capa_Negocio_Eliminar_Tabla_Temporal_Actualizar_Fecha(int id_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Eliminar_Tabla_Temporal_Actualizar_Fecha(id_usuario);
            }
            catch (Exception e)
            {

                throw e;
            }

        }



        public List<Cls_Entidad_Importacion_Lecturas> Capa_Negocio_Registros_Incorrectos(int idlocal, int idservicio, int iduser)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Registros_Incorrectos(idlocal, idservicio, iduser);
            }
            catch (Exception e)
            {

                throw e;
            }
        }
        
        public bool Capa_Negocio_Guardando_Informacion_Actualizacion_Fecha_Lectura(int id_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Guardando_Informacion_Actualizacion_Fecha_Lectura(id_usuario);
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public string Capa_Negocio_MigracionTemporalLectura(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> listaobjeto = new List<Cls_Entidad_Lecturas_Relecturas>();
                
                bool valor = true;
                string resultado = "";
  
                Cls_Dato_Importacion_Lecturas Objeto_Dato01 = new Cls_Dato_Importacion_Lecturas();
                resultado = Objeto_Dato01.Capa_Dato_InsertandoUpdate_Lecturas(fechaAsignacion, usuario);
                
                return resultado;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public string Capa_Negocio_save_Lecturas(string fechaAsignacion , string fechaMovil, int id_servicio, string nombre_archivo, int usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> listaobjeto = new List<Cls_Entidad_Lecturas_Relecturas>();
               string resultado = "";
                Cls_Dato_Importacion_Lecturas Objeto_Dato01 = new Cls_Dato_Importacion_Lecturas();
                resultado = Objeto_Dato01.Capa_Dato_save_Lecturas(fechaAsignacion, fechaMovil, usuario);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public string Capa_Negocio_save_grandesClientes(string fechaAsignacion, string fechaMovil, int id_servicio, string nombre_archivo, int usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> listaobjeto = new List<Cls_Entidad_Lecturas_Relecturas>();
                string resultado = "";
                Cls_Dato_Importacion_Lecturas obj = new Cls_Dato_Importacion_Lecturas();
                resultado = obj.Capa_Dato_save_grandesClientes(fechaAsignacion, fechaMovil, usuario);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public string Capa_Negocio_save_Lecturas_Reclamos(string fechaAsignacion, string fechaMovil, int id_servicio, string nombre_archivo, int usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> listaobjeto = new List<Cls_Entidad_Lecturas_Relecturas>();
                string resultado = "";
                Cls_Dato_Importacion_Lecturas Objeto_Dato01 = new Cls_Dato_Importacion_Lecturas();
                resultado = Objeto_Dato01.Capa_Dato_save_Lecturas_Reclamos(fechaAsignacion, fechaMovil, usuario);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public object Capa_Negocio_generarRepartoPDf(string fechaAsignacion, int tipo)
        {
            try
            {
               Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
               return Objeto_Dato.Capa_Dato_get_generarReparto_Pdf(fechaAsignacion, tipo);
             }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public object Capa_Negocio_generarRepartoPDf_individual(string fechaAsignacion, string suministro)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_get_generarReparto_Pdf_individual(fechaAsignacion, suministro);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public string Capa_Negocio_save_Reclamos(string fechaAsignacion, string fechaMovil, int id_servicio, string nombre_archivo, int usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> listaobjeto = new List<Cls_Entidad_Lecturas_Relecturas>();
                string resultado = "";
                Cls_Dato_Importacion_Lecturas Objeto_Dato01 = new Cls_Dato_Importacion_Lecturas();
                resultado = Objeto_Dato01.Capa_Dato_save_Reclamos(fechaAsignacion, fechaMovil, usuario);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public string Capa_Negocio_MigracionTemporalRelectura(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario)
        {
           Cls_Dato_Importacion_Lecturas oClsNegocioImportacionLecturas = new Cls_Dato_Importacion_Lecturas();
            return oClsNegocioImportacionLecturas.Capa_Dato_MigracionTemporalLectura(fechaAsignacion, id_servicio, nombre_archivo, usuario);
        }

        
        public bool Capa_Negocio_ListaExcel(string fileLocation, int usuario, int idlocal, string idfechaAsignacion)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_ListaExcel(fileLocation, usuario, idlocal, idfechaAsignacion);
            }
            catch (Exception e)
            {

                throw e;
            }

        }


        public bool Capa_Negocio_Eliminar_Tabla_Temporal(int id_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Eliminar_Tabla_Temporal(id_usuario);
            }
            catch (Exception e)
            {

                throw e;
            }

        }


        public bool Capa_Negocio_Guardar_Informacion(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Guardar_Informacion(fechaAsignacion, id_servicio, nombre_archivo, usuario);
            }
            catch (Exception)
            {

                throw;
            }

        }



        public List<Servicio> Capa_Negocio_PermisoListUsuarioServicio(int idusuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_ListarPermisoUsuarioServicio(idusuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_get_marcaMedidor()
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_get_marcaMedidor();
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public object Capa_Negocio_buscarCodigoEMr(string codigo, string fechaCarga)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_get_buscarCodgioEmr(codigo, fechaCarga);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public int Capa_Negocio_grabarGrandesClienteFile(int Id_GrandeCliente, string CodigoEMR, string nameFile_GrandeClienteFile, string urlNameFile_GrandeClienteFile, int id_marcaMedidor , string fechaCarga, int idUsuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_grabarGrandesClienteFile(Id_GrandeCliente, CodigoEMR, nameFile_GrandeClienteFile, urlNameFile_GrandeClienteFile, id_marcaMedidor, fechaCarga, idUsuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<Servicio> Capa_Negocio_Listado_Servicios()
        {

            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Listado_Servicios();
            }
            catch (Exception e)
            {

                throw e;
            }


        }


        public bool Capa_Negocio_ImportarArchivoExcel(Cls_Entidad_Importacion_Lecturas Entidad)
        {

            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_ImportarArchivoExcel(Entidad);
            }
            catch (Exception e)
            {

                throw e;
            }


        }

        public List<Cls_Entidad_Importacion_Lecturas> Capa_Negocio_Listar_Temporal_Lecturas(int cod_usuario)
        {

            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Listar_Temporal_Lecturas(cod_usuario);
            }
            catch (Exception e)
            {

                throw e;
            }


        }


        public List<Cls_Entidad_Importacion_Lecturas> Capa_Negocio_Listar_Temporal_Lecturas_Agrupado(int cod_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Listar_Temporal_Lecturas_Agrupado(cod_usuario);
            }
            catch (Exception e)
            {

                throw e;
            }


        }



        public List<Cls_Entidad_Importacion_Lecturas> Capa_Negocio_Listar_Temporal_Lecturas_Detallado(int cod_usuario, string SC, string DNI)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Listar_Temporal_Lecturas_Detallado(cod_usuario, SC, DNI);
            }
            catch (Exception e)
            {

                throw e;
            }


        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <param name="usuario"></param>
        /// <param name="idlocal"></param>
        /// <param name="idservicio"></param>
        /// <param name="idfechaAsignacion"></param>
        /// <returns></returns>
        public bool Capa_Negocio_TemporalLecturaCargar(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            bool res = false;
            Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
            res = Objeto_Dato.Capa_Dato_TemporalLectura(fileLocation, usuario, idlocal, idservicio, idfechaAsignacion, nombreArchivo);
            return res;
        }


        public object Capa_Negocio_save_temporalLectura(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_save_temporalLectura(fileLocation, usuario, idlocal, idservicio, idfechaAsignacion, nombreArchivo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_save_Temporal_grandesClientes(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_save_Temporal_grandesClientes(fileLocation, usuario, idlocal, idservicio, idfechaAsignacion, nombreArchivo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_save_temporalLectura_reclamos(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_save_temporalLectura_Reclamos(fileLocation, usuario, idlocal, idservicio, idfechaAsignacion, nombreArchivo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_save_temporalReclamos(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_save_temporalReclamos(fileLocation, usuario, idlocal, idservicio, idfechaAsignacion, nombreArchivo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Negocio_Listar_Detalle_Lectura_Relectura(int cod_usuario, int idtecnico, string distrito)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Listar_TemporalLectura_Relectura(cod_usuario, idtecnico, distrito);
            }
            catch (Exception e)
            {

                throw e;
            }


        }


        /// <summary>
        /// Lista  de datos lecturas y relecturas
        /// </summary>
        /// <param name="cod_usuario"></param>
        /// <returns></returns>
        public object  Capa_Negocio_Agrupado_temporal_lectura(string idfechaAsignacion, int cod_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Agrupado_lectura(idfechaAsignacion, cod_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public object Capa_Negocio_Agrupado_temporal_grandesClientes(string idfechaAsignacion, int cod_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Agrupado_grandesClientes(idfechaAsignacion, cod_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public object Capa_Negocio_Agrupado_temporal_lecturaReclamo(string idfechaAsignacion, int cod_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Agrupado_lectura_reclamos(idfechaAsignacion, cod_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public object Capa_Negocio_Agrupado_temporal_reclamos(string idfechaAsignacion, int cod_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Agrupado_reclamos(idfechaAsignacion, cod_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
 



        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Negocio_Agrupado_temporal_Relectura(string idfechaAsignacion, int cod_usuario)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Agrupado_temporal_Relectura(idfechaAsignacion, cod_usuario);
            }
            catch (Exception e)
            {

                throw e;
            }
        }


        public string Capa_Negocio_set_EnviarCorreo(int servicio, string fecha_Asigna,int usuario, int local)
        {
            try
            {
                string resultado = "";
                Cls_Dato_Importacion_Lecturas Objeto_Dato01 = new Cls_Dato_Importacion_Lecturas();
                resultado = Objeto_Dato01.Capa_Dato_set_EnviarCorreo(servicio, fecha_Asigna, usuario, local);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public object Capa_Negocio_get_repartoPDF(int servicio, string fecha_Asigna)
        {
            try
            {
                Cls_Dato_Importacion_Lecturas Objeto_Dato01 = new Cls_Dato_Importacion_Lecturas();
                return Objeto_Dato01.Capa_Dato_get_repartoPDF(servicio, fecha_Asigna); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}
