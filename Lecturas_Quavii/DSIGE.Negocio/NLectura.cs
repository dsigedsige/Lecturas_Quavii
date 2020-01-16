using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Dato;
using System.Data;

namespace DSIGE.Negocio
{
    public class NLectura
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-03
        /// </summary>
        public NLectura() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-03
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        /// 



        public DataTable ListaExcel(string fileLocation)
        {
            return new DImportarDiasTrabajo().ListaExcel(fileLocation);
        }

        //public bool NImportarArchivoExcel01(List<LeerLecturas> ListLectura)
        //{
        //    return new DLectura().ImportarArchivoExcel(ListLectura);
        //}


        public bool NImportarArchivoExcel(List<Lectura> ListLectura)
        {
            return new DLectura().ImportarArchivoExcel(ListLectura);
        }

        public List<Lectura> NListarTablaTem(int idemp)
        {
            return new DLectura().DListarTablaTem(idemp);
        }

        public bool Guardar_TableTemp_Lecturas(int idUsu, string fechaAsignacion)
        {
            return new DLectura().Guardar_TableTemp_Lecturas(idUsu, fechaAsignacion);
        }


        public bool InsertarExcel(int idlocal, int idUsu, DataTable dt)
        {

            return new DLectura().InsertarExcel(idlocal, idUsu,dt);
        }

        public int NBulkInfo(List<Lectura> oRq)
        {
            try {
                return new DLectura().DBulkInfo(oRq);
            }
            catch (Exception e) {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-07
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Lectura> NListaValidado(Request_Lectura_Validacion oRq)
        {
            try {
                return new DLectura().DListaValidado(oRq);
            }
            catch (Exception e) {
                throw e;
            }
        }


        public List<Cls_Entidad_ReporteDiarioNew> Capa_Negocio_Get_ListadoReporteDiario(int cliente, string fechaini, string fechafin, int idServicio)
        {
            try
            {
                DLectura Objeto_Dato = new DLectura();
                return Objeto_Dato.Capa_Dato_Get_ListadoReporteDiario(cliente, fechaini, fechafin, idServicio);
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-08
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NInsertaMasivo(Request_Lectura_Inserta_Masivo oRq)
        {
            try
            {
                return new DLectura().DInsertaMasivo(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Fotos> Capa_Negocio_Get_MostrarFotos(int id_lectura, int idservicio)
        {
            try
            {
                DLectura Objeto_Dato = new DLectura();
                return Objeto_Dato.Capa_Dato_Get_MostrarFotos(id_lectura, idservicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<ResumenLecturasCab> NListaLecturaAgrupado(int id_local, int id_servicio, string fechaini, string fechaFin)
        {
            try
            {
                return new DLectura().DListaLecturaAgrupado(id_local, id_servicio, fechaini, fechaFin);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ResumenLecturasCab> NListaLecturaAgrupado_new(int id_local, int id_servicio, string fechaini, string fechaFin, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                return new DLectura().DListaLecturaAgrupado_new(id_local, id_servicio, fechaini, fechaFin, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ResumenLecturas> NListaLecturaDetallado(int id_local, int id_servicio, string fechaini, string fechaFin, int idoperario)
        {
            try
            {
                return new DLectura().DListaLecturaDetallada(id_local, id_servicio, fechaini, fechaFin, idoperario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<ResumenLecturas> NListaLecturaDetalladoEstado(int id_local, int id_servicio, string fechaini, string fechaFin, int idoperario, int categoriaLectura)
        {
            try
            {
                return new DLectura().DListaLecturaDetalladaEstado(id_local, id_servicio, fechaini, fechaFin, idoperario, categoriaLectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ResumenLecturas> Capa_Negocio_Get_MostrarFotosLecturas(int id_lectura)
        {
            try
            {
                DLectura Objeto_Dato = new DLectura();
                return Objeto_Dato.Capa_Dato_Get_MostrarFotosLecturas(id_lectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaHistorico> NListaLecturasResumen(Request_Lectura_Historico oRq)
        {
            try
            {
                return new DLectura().DListaLecturaResumen(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaHistorico> NListaLecturasResumenObs(Request_Lectura_Historico oRq, string lista)
        {
            try
            {
                return new DLectura().DListaLecturaResumenObs(oRq, lista);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaTomadas> NListaLecturaTomadaFoto(Request_Lectura_Tomadas oRq, string lista, int id_operario)
        {
            try
            {
                return new DLectura().DListaLecturaTomadaFoto(oRq, lista, id_operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public List<urlLectura> NListaUrlLectura(string id,string fechaini, string fechafin,int con)
        {
            try
            {
                return new DLectura().DListaUrlLectura(id,fechaini,fechafin,con);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaTomadas> NListaLecturaHistoricoXSuministro(Request_Lectura_Tomadas oRq)
        {
            try
            {
                return new DLectura().DListaLecturaHistoricoXSuministro(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaTomadas> NListaFotosMasivo()
        {
            try
            {
                return new DLectura().DListaFotosMasivo();
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Lecturas_Tomadas> NListaLecturasTomadas(Request_Lectura_Tomadas oRq)
        {
            try
            {
                return new DLectura().DListaLecturasTomadas(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public DataTable N_ListandoReparto_Tomadas(int servicio, string tipoRecibo, string cicloFacturacion, int Estado, string fecha_ini, string fecha_fin, string suministro, string medidor, int operario)
        {
            try
            {
                return new DLectura().D_ListandoReparto_Tomadas(servicio, tipoRecibo, cicloFacturacion, Estado, fecha_ini, fecha_fin, suministro, medidor, operario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Lecturas_Tomadas> N_Descarga_ResultadosReclamosExcel(string fechaini, string fechafin, int servicio)
        {
            try
            {
                return new DLectura().D_Descarga_ResultadosReclamosExcel(fechaini, fechafin, servicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        


        public bool NImportarArchivoExcel01(List<LeerLecturas> ListLectura)
        {
            return true;
        }

 

    }
}
