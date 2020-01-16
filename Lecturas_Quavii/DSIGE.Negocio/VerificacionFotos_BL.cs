using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Modelo
{
    public class VerificacionFotos_BL
    {

        public List<VerificacionFoto_E> Capa_Negocio_Get_ListaServicios()
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<HistoricoLectura> Capa_Negocio_Get_ListaHistoricoLectura(string suministro)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaHistoricoLectura(suministro);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<VerificacionFoto_E> Capa_Negocio_Get_ListaOperarios()
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaOperarios();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<VerificacionFoto_E> Capa_Negocio_Get_ListaSector()
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaSector();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<VerificacionFoto_E> Capa_Negocio_Get_ListaObservacion()
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaObservacion();
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<VerificacionFoto_E> Capa_Negocio_Get_ListaFotoLecturas(int local, string fecha, int servicio, int operario, int observacion, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaFotoLecturas(local, fecha, servicio, operario, observacion, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<VerificacionFoto_E> Capa_Negocio_Get_ListaFotoLecturas_SinObs(string fecha, int servicio, int operario, int observacion, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaFotoLecturas_SinObs(fecha, servicio, operario, observacion, id_supervisor, id_operario_supervisor);
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
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Set_VerificacionFotos(lecturas);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_VerificacionLecturas_Inconsistencias(string lecturas)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Set_VerificacionFotos_Inconsistencias(lecturas);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_EnviarFotosSinObservacion(string lecturas)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_EnviarFotosSinObservacion(lecturas);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public string Capa_Negocio_CambiarLectura(int id_lectura, string confirmacion_lectura)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_CambiarLectura(id_lectura, confirmacion_lectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public List<VerificacionFoto_E> Capa_Negocio_Get_ListaFotoLecturasValidacion(string fecha, int servicio, int operario, string sector, int observacion)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_ListaFotoLecturasValidacion(fecha, servicio, operario, sector, observacion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_Get_Update_LecturaActual(int id_lectura, int id_usuario, string lectura_actual)
        {
            Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
            return Objeto_Dato.Capa_Dato_Get_Update_LecturaActual(id_lectura, id_usuario, lectura_actual);

        }

        public List<VerificacionFoto_E> Capa_Negocio_Get_MostrarFotos(int id_lectura)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_MostrarFotos(id_lectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<VerificacionFoto_E> Capa_Negocio_Get_MostrarFotos_Reparto(int idReparto)
        {
            try
            {
                Cls_Dato_VerificacionFotos Objeto_Dato = new Cls_Dato_VerificacionFotos();
                return Objeto_Dato.Capa_Dato_Get_MostrarFotos_Reparto(idReparto);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        

    }
}
