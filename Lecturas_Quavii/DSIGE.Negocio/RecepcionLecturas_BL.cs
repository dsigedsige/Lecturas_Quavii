using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;

namespace DSIGE.Modelo
{
    public class RecepcionLecturas_BL
    {

        public List<Cls_Entidad_RecepcionLect> Capa_Negocio_Get_ListaServicios()
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Negocio_Get_ListaAlejadosCorLdsNot(int id_servicio, string id_sector, string fecha_asignacion)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_GerReporteFueraRangoCorLdsNot(id_servicio, id_sector, fecha_asignacion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Negocio_Get_ListaAlejadosCorLds(int id_servicio, string id_sector, string fecha_asignacion, int tipo)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_GerReporteFueraRangoCorLds(id_servicio, id_sector, fecha_asignacion, tipo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Negocio_Get_ListaAlejadosCorLds_Det(int id_servicio, string id_sector, string fecha_asignacion, string zona_lectura, int tipo)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_GerReporteFueraRangoCorLds_Det(id_servicio, id_sector, fecha_asignacion, zona_lectura, tipo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Negocio_Get_ListaAlejadosCorLds_DetNot(int id_servicio, string id_sector, string fecha_asignacion, int id_operario)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_GerReporteFueraRangoCorLds_DetNot(id_servicio, id_sector, fecha_asignacion, id_operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Negocio_Get_ListaAlejadosCorLds_DetObs(int id_servicio, string id_sector, string fecha_asignacion, string zona_lectura, int tipo)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_GerReporteFueraRangoCorLds_DetObs(id_servicio, id_sector, fecha_asignacion, zona_lectura, tipo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Lecturas_Dia> Capa_Negocio_Get_ListadoLecturasDia(int id_local, string fecha, int id_servicio, int tipoProceso, int pendiente, int asignado, int verificado)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListadoLecturasDia(id_local, fecha, id_servicio, tipoProceso, pendiente, asignado, verificado);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<Cls_Entidad_Lectura_OperarioZona> Capa_Negocio_Get_ListadoOperariosZona(int id_local, string fecha, int id_servicio, int tipoProceso)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListadoOperariosZona(id_local, fecha, id_servicio, tipoProceso);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public object Capa_Negocios_GuardarTemporalSum(List<Cls_Entidad_Temporal_Sum> listLecturas, int tipoServicio, string fecha_asignacion)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_GuardarTemporalSuministro(listLecturas, tipoServicio, fecha_asignacion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocios_AnularLectura(int id_lectura)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_AnularLectura(id_lectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocios_GenerarSecuencia(int idLecturista, int idServicios, string idSector, string FechaInicial, int Ano, int Mes)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_GenerarSecuencia(idLecturista, idServicios, idSector, FechaInicial, Ano, Mes);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Resumen_Personal> Capa_Negocio_Get_ListadoResumenOperario(string fecha, int id_Operario)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListadoResumenOperario(fecha, id_Operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<Cls_Entidad_RecepcionLect> Capa_Negocio_Get_ListaOperarios()
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListaOperario();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Lecturas_Trabajo> Capa_Negocio_Get_ListaTrabajosResumen(string fecha, int id_servicio)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListaTrabjosResumen(fecha, id_servicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Lecturas_Trabajo> Capa_Negocio_Get_ListaTrabajosDetalle(string fecha, int id_servicio)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListaTrabjosDetalle(fecha, id_servicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public List<Cls_Entidad_RecepcionLect> Capa_Negocio_Get_ListaEstados()
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListandoEstados();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_RecepcionLect> Capa_Negocio_Get_ListaInformacionLecturas(int local, string fechaAsigna, int servicio, int id_operador, string rt, string s1, string sector, string zona, string suministro, string medidor, int estado)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListarInformacionLecturas(local, fechaAsigna, servicio, id_operador, rt, s1, sector, zona, suministro, medidor, estado);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_RecepcionLect> Capa_Negocio_Get_ListaInformacionLecturas_new(int local, string fechaAsigna, int servicio, int id_operador, string rt, string s1, string sector, string zona, string suministro, string medidor, int estado, int id_zona)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Dato_Get_ListarInformacionLecturas_new(local, fechaAsigna, servicio, id_operador, rt, s1, sector, zona, suministro, medidor, estado, id_zona);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_RecepcionLect> Capa_Negocio_Get_ListaInformacionLecturas_Bak(string fecha_movil, string fecha_servidor)
        {
            try
            {
                Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
                return Objeto_Dato.Capa_Negocio_Get_ListaInformacionLecturas_Bak(fecha_movil, fecha_servidor);
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        public string Capa_Negocio_Get_Update_Lectura(int id_lectura, string direccion, string lectura_actual, string observacion, string medidor, int id_usuario, string dns, string ubicacion)
        {
            Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
            return Objeto_Dato.Capa_Dato_Update_Lectura(id_lectura, direccion, lectura_actual, observacion, medidor, id_usuario, dns, ubicacion);

        }

        public string Capa_Negocio_Update_operarios_Historial(string listIds)
        {
            Cls_Dato_RecepcionLect Objeto_Dato = new Cls_Dato_RecepcionLect();
            return Objeto_Dato.Capa_Dato_Update_Historial_Operarios(listIds);

        }



    }
}
