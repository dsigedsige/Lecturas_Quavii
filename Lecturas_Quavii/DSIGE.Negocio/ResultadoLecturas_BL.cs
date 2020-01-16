using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class ResultadoLecturas_BL
    {

        public List<ResultadoLecturas_E> Capa_Negocio_Get_ListaServicios()
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ResultadoLecturas_E> Capa_Negocio_Get_ListaOperarios()
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_ListaOperarios();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ResultadoLecturas_E> Capa_Negocio_Get_ListaSector()
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_ListaSector();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ResultadoLecturas_E> Capa_Negocio_Get_ListaResumenLectura(int Local , string FechaAsignacion, int id_tiposervicio, int id_supervisor, int id_operario_supervisor, string ciclo)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_ResumenLecturas(Local, FechaAsignacion, id_tiposervicio, id_supervisor, id_operario_supervisor, ciclo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<ResultadoLecturas_E> Capa_Negocio_Get_ListaLectura_Observacion(int anio, int mes, string sector)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_ResumenLecturas_Observacion(anio, mes, sector);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<ResultadoLecturas_E> Capa_Negocio_Get_ListaLectura_Detallado(int Local, string FechaAsignacion, int id_tiposervicio, int id_supervisor, int id_operario_supervisor,  string ciclo)           
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_ResumenLecturas_Detallado(Local, FechaAsignacion, id_tiposervicio, id_supervisor, id_operario_supervisor, ciclo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_VerificacionLecturas(string lecturas)

        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Set_VerificacionFotos(lecturas);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<ResultadoLecturas_E> Capa_Negocio_Get_FotoSelfie(int anio, int mes, string sector, int operario)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_FotoSelfie(anio, mes, sector, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<ResultadoLecturas_E> Capa_Negocio_Get_NotasOperario(string fechaAsignacion, int idServicio, int operario)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_NotasOperario(fechaAsignacion, idServicio, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ResultadoLecturas_E> Capa_Negocio_Get_Observaciones(string fechaAsignacion, int idServicio, int operario, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_Observaciones(fechaAsignacion, idServicio, operario, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_Guardar_NotasOperario(string fechaAsignacion, int idServicio, int operario, string observacion, int usuario)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Guardar_NotasOperario(fechaAsignacion, idServicio,  operario, observacion, usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<ResultadoLecturas_E> Capa_Negocio_Get_ListaUltimoSector()
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Get_ultimoSector();
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        public string Capa_Negocio_Descargar_efectividad( int Local, int servicio, string fecha, int id_usuario, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Descargar_efectividad(Local, servicio, fecha, id_usuario, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public string Capa_Negocio_Descargar_efectividad_detallado(int Local,  int servicio, string fecha, int id_usuario)
        {
            try
            {
                Cls_Dato_ResultadoLecturas Objeto_Dato = new Cls_Dato_ResultadoLecturas();
                return Objeto_Dato.Capa_Dato_Descargar_efectividad_detallado(Local, servicio, fecha, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


    }
}
