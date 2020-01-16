using DSIGE.Dato;
using DSIGE.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Negocio
{
    public class NUbicacion_Lecturas
    {
        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <returns></returns>
        public List<UbicacionOperario> NListaLecturasOperariosGPS(int idOpe, string fechaAsig, string suministro, string medidor, int idEmp, string lista)
        {
            try
            {
                return new DUbicacion_Lecturas().DListaLecturasOperariosGPS(idOpe, fechaAsig, suministro, medidor, idEmp, lista);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<Ubicacion_Lectura> Capa_Negocio_Get_ListaOperarios()
        {
            try
            {
                DUbicacion_Lecturas Objeto_Dato = new DUbicacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Get_ListaOperarios();
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Ubicacion_Lectura> Capa_Negocio_Get_ListaServicios()
        {
            try
            {
                DUbicacion_Lecturas Objeto_Dato = new DUbicacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<UbicacionOperario> NSeguimiento_Lecturas_II(int Local, string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            try
            {
                return new DUbicacion_Lecturas().DSeguimiento_Lectura_II(Local, pFechaAsiga, servicio, operario, suministro, medidor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public List<UbicacionOperario> NSeguimiento_Lecturas_III(string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            try
            {
                return new DUbicacion_Lecturas().DSeguimiento_Lectura_III(pFechaAsiga, servicio, operario, suministro, medidor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<UbicacionOperario> NSeguimiento_Lecturas_Resumen(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                return new DUbicacion_Lecturas().DSeguimiento_Lectura_Resumen(FechaAsiga, servicio, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Ubicacion_Lectura> Capa_Negocio_Get_UbicacionLecturas(int Local, string fechaAsigna, int servicio, int operario)
        {
            try
            {
                DUbicacion_Lecturas Objeto_Dato = new DUbicacion_Lecturas();
                return Objeto_Dato.Capa_Dato_Get_UbicacionLecturas(Local, fechaAsigna, servicio, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
