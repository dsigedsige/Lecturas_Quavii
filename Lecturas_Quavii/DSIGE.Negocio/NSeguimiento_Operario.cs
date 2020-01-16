using DSIGE.Dato;
using DSIGE.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Negocio
{
    public class NSeguimiento_Operario
    {

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 01-10-2015
        /// </summary>
        /// <returns></returns>
        //public List<UbicacionOperario> NSeguimientoOperariosGPS(int idOperario)
        //{
        //    try
        //    {
        //        return new DSeguimiento_Operario().DSeguimientoOperarioGPS(idOperario);
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}

        //public List<UbicacionOperario_GPS> NSeguimientoOperario_GPS(int pIdOpe, string  pFechaAsiga, string lista)
        //{
        //    try
        //    {
        //        return new DSeguimiento_Operario().DSeguimientoOperario_GPS(pIdOpe, pFechaAsiga,lista);
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}

        //public List<UbicacionOperario_GPS> NSeguimientoOperario_GPS_II(string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        //{
        //    try
        //    {
        //        return new DSeguimiento_Operario().DSeguimientoOperario_GPS_II(pFechaAsiga, servicio, operario, suministro, medidor);
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}


        public List<UbicacionOperario> NSeguimientoOperariosGPS(int idOperario)
        {
            try
            {
                return new DSeguimiento_Operario().DSeguimientoOperarioGPS(idOperario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<UbicacionOperario_GPS> NSeguimientoOperario_GPS(int pIdOpe, string pFechaAsiga, string lista)
        {
            try
            {
                return new DSeguimiento_Operario().DSeguimientoOperario_GPS(pIdOpe, pFechaAsiga, lista);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<UbicacionOperario_GPS> NSeguimientoOperario_GPS_II(int Local , string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            try
            {
                return new DSeguimiento_Operario().DSeguimientoOperario_GPS_II(Local, pFechaAsiga, servicio, operario, suministro, medidor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<UbicacionOperario_GPS> NSeguimientoOperario_GPS_reparto(string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            try
            {
                return new DSeguimiento_Operario().DSeguimientoOperario_GPS_Reparto(pFechaAsiga, servicio, operario, suministro, medidor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<UbicacionOperario_GPS> NSeguimientoOperario_GPS_Resumen(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                return new DSeguimiento_Operario().DSeguimientoOperario_GPS_Resumen(FechaAsiga, servicio, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Ndescargar_informacion_reparto_excel(string pFechaAsiga, int servicio, int operario, string suministro, string medidor, int id_usuario)
        {
            try
            {
                return new DSeguimiento_Operario().Ddescargar_informacion_reparto_excel(pFechaAsiga, servicio, operario, suministro, medidor, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_get_Operario_Gps(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                return new DSeguimiento_Operario().Capa_Dato_get_Operario_Gps(FechaAsiga, servicio, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public object Capa_Negocio_get_Suministros_Mapa(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                return new DSeguimiento_Operario().Capa_Dato_get_Suministros_Mapa(FechaAsiga, servicio, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public object Capa_Negocio_get_Suministros_Trabajados(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                return new DSeguimiento_Operario().Capa_Datos_get_Suministros_Trabajados(FechaAsiga, servicio, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


    }
}
