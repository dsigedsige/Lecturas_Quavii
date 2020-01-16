using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
    public class Cls_Negocio_AsignarOrdenTrabajo
    {
        //Listado

        public List<Cls_Entidad_AsignaOrdenTrabajo.Locales> Capa_Negocio_Get_ListaLocales()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaLocales();
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaServicio()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio();
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        /// <summary>
        /// Lista servicios por usuario
        /// </summary>
        /// <returns></returns>
        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaServicioXusuario(int idusuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaServicioXusuario(idusuario);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaServicioXusuario_II(int idusuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaServicioXusuario_II(idusuario);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaServicio_new()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio_new();
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.Estados> Capa_Negocio_Get_ListaEstados()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaEstados();
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        public object ListaPreEnvioExcelTexto(int id_local,int id_tipo_servicio,int estado,string fechaAsignacion)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaPreEnvioExcelTexto(id_local, id_tipo_servicio, estado, fechaAsignacion);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Estados> Capa_Negocio_Get_ListaEstadosLectura()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaEstadosLecturaEnviado();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Estados> Capa_Negocio_Get_estadosAll()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_estadosAll();
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        


        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Negocio_Get_ListaObservaciones(int TipoServicio)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaObservaciones(TipoServicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_set_anulandoFoto(int id_lectura, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_set_anulandoFoto(id_lectura, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_getGrandesClientes(int estado, string fecha_inicial, string fecha_final, string codigoEmr)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_get_grandesClientes(estado, fecha_inicial, fecha_final, codigoEmr);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
                       

        public object Capa_Negocio_get_ultimoCodigoEmr(int idusuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_get_ultimoCodigoEmr(idusuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_download_grandesClientes(int estado, string fecha_inicial, string fecha_final, string codigoEmr, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_download_grandesClientes(estado, fecha_inicial, fecha_final, codigoEmr, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Negocio_getGrandesClientes_detalle(int Id_GrandeCliente)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_get_grandesClientes_detalle(Id_GrandeCliente);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public object Capa_negocio_get_download_grandesClientes(int Id_GrandeCliente, int tipo, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_get_download_grandesClientes(Id_GrandeCliente, tipo, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        




        public object Capa_Negocio_getGrandesClientes_detalleFile(int Id_GrandeCliente)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_get_grandesClientes_detalleFile(Id_GrandeCliente);
            }
            catch (Exception e)
            {
                throw e;
            }
        }




        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Negocio_Get_ListaObservacionesLectura()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ObservacionesLecturas();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Negocio_Get_ListaObservacionesCortes()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ObservacionesCortes();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Negocio_Get_ListaObservacionesCortesResultado()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ObservacionesCortesResultado();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Negocio_Get_ListaObservacionesReconexion()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ObservacionesReconexion();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Negocio_Get_ListaObservacionesReconexionResultado()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ObservacionesReconexionResultado();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
                                      
        public string Capa_Negocio_Set_ActualizandoLecturasRelecturas(int cod_lectura, string DatoModificar, string campoModificar)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Set_ActualizandoLecturasRelecturas(cod_lectura,DatoModificar, campoModificar);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public string Capa_Negocio_Set_ActualizandoCortesReconexiones(int cod_Cortelectura, string DatoModificar, string campoModificar)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Set_ActualizandoCortesReconexiones(cod_Cortelectura, DatoModificar, campoModificar);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.Tecnico> Capa_Negocio_Get_ListaTecnicos(int id_local, int id_tipo_servicio, int id_empresa)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaTecnico(id_local, id_tipo_servicio, id_empresa);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.OrdenesTrabajo> Capa_Negocio_Get_ListandoOrdenesTrabajoGeneral(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion, string horainicial, string horafinal)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListandoOrdenesTrabajoGeneral(empresa,id_local, id_tipo_servicio, estado,suministro, medidor, tecnico, fechaAsignacion, horainicial , horafinal );
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Negocio_Get_ListaLecturaEnviarCliente(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, 
                                                    int tecnico, string fechaAsignacion, string tipoCliente, int id_supervisor, int id_operario_supervisor, string horainicial, string horafinal)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaLecturaEnviarCliente(empresa, id_local, id_tipo_servicio, estado, suministro, medidor, tecnico, fechaAsignacion, tipoCliente, id_supervisor, id_operario_supervisor, horainicial , horafinal );
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Negocio_Get_ListaLectura_EnviarCliente(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion, string tipoCliente, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaLectura_EnviarCliente(empresa, id_local, id_tipo_servicio, estado, suministro, medidor, tecnico, fechaAsignacion, tipoCliente, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public string Capa_Negocio_Get_DescargaFoto(string List_codigos, int servicio, int flag_multiple, int usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_DescargaFoto(List_codigos, servicio, flag_multiple, usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_Get_DescargaFoto_sinObservacion(string List_codigos, int servicio, int flag_multiple, int usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_DescargaFoto_sinObservacion(List_codigos, servicio, flag_multiple, usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public string Capa_Negocio_Get_DescargaFoto_new(int local,  int servicio, string fecha_asignacion, int id_usuario, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_DescargaFoto_new(local, servicio, fecha_asignacion, id_usuario, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_Get_DescargaFoto_new_sinObservacion(int local, int servicio, string fecha_asignacion, int id_usuario, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_DescargaFoto_new_sinObservacion(local, servicio, fecha_asignacion, id_usuario, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Negocio_DescargaExcel_PreEnvio(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_DescargaExcel_PreEnvio(empresa, id_local, id_tipo_servicio, estado, suministro, medidor, tecnico, fechaAsignacion, id_supervisor, id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Negocio_DescargaExcel_EnvioCliente(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_DescargaExcel_EnvioCliente(empresa, id_local, id_tipo_servicio, estado, suministro, medidor, tecnico, fechaAsignacion);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public string Capa_Negocio_InsertandoUpdate_LecturaMeses(int user, string fechaAsigna)
        {
            try
            {
                Cls_Dato_Export_trabajos_lectura Objeto_Dato = new Cls_Dato_Export_trabajos_lectura();
                return Objeto_Dato.Capa_Dato_InsertandoUpdate_LecturaMeses(user, fechaAsigna);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> NListaFotos(int idLectura, int opcion)
        {
            try
            { 
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.DListaFotos(idLectura, opcion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura> Capa_Negocio_LecturasRelectura_Historico(int empresa, string suministro, int tiposervicio)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_LecturasRelectura_Historico(empresa, suministro, tiposervicio);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Historico_CortesReconexiones> Capa_Negocio_CortesReconexiones_Historico(int empresa, string suministro, int tiposervicio)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_CortesReconexiones_Historico(empresa, suministro, tiposervicio);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public string Capa_Negocio_UpdateFechaAplicativoMovil(int empresa, string fechamovil, int id_usuario, int tipoServicio, List<int> List_codigos)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_UpdateFechaAplicativoMovil(empresa, fechamovil, id_usuario, tipoServicio, List_codigos);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public string Capa_Negocio_UpdateFechaAplicativoMovilLectura(int empresa, string fechamovil, int id_usuario, int tipoServicio, List<int> List_codigos)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_UpdateFechaAplicativoMovilLectura(empresa, fechamovil, id_usuario, tipoServicio, List_codigos);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_UpdateReasignaOperador(int empresa, int usuario, int TipoServicio, int tecnico, string fechaReasigna, List<int> List_codigos, int opcion)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_UpdateReasignaOperador(empresa, usuario, TipoServicio, tecnico, fechaReasigna, List_codigos, opcion);
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        
        public List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> Capa_Negocio_Get_ListaVisualizacionesPendientes(int local, string fechaAsigna, int TipoServicio)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaVisualizacionesPendientes(local, fechaAsigna, TipoServicio);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> Capa_Negocio_Get_ListaVisualizacionesPendientes_Detallado(int TipoServicio, string id_operario, string  id_estado, string fechaAsigna)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaVisualizacionesPendientes_Detallado(TipoServicio, id_operario, id_estado, fechaAsigna);
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        
        public List<Cls_Entidad_AsignaOrdenTrabajo.validacion> Capa_Negocio_Get_ListaLecturaEnviarCliente_Validacion(int estado, string fechaAsignacion, string estado_new, int resultado, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaLecturaEnviarCliente_Validacion( estado, fechaAsignacion, estado_new,resultado, id_supervisor , id_operario_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        
        public string Capa_negocio_DescargarArchivoTexto_EnvioCliente(int local, int servicio, int estado, string fechaAsigna, int id_supervisor, int id_operario_supervisor)
        {
            Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
            return Objeto_Dato.Capa_dato_DescargarArchivoTexto_EnvioCliente(local, servicio, estado, fechaAsigna, id_supervisor, id_operario_supervisor);
        }

        public string Capa_negocio_DescargarArchivoTexto_CortesReconexiones(int servicio, int estado, string fechaAsigna, int id_operario, int id_usuario, int tipo, string horainicial, string horafinal)
        {
            Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
            return Objeto_Dato.Capa_dato_DescargarArchivoTexto_cortesReconexiones(servicio, estado, fechaAsigna, id_operario, id_usuario, tipo, horainicial, horafinal);
        }


        public string Capa_Negocio_Set_procesarRecepcionLecturas(int servicio, string fechaAsigna)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Set_procesarRecepcionLecturas(servicio, fechaAsigna);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_Set_procesarRecepcion_Trabajos(int servicio, string fechaAsigna)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Set_procesarRecepcion_Trabajos(servicio, fechaAsigna);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_Proceso_Verificacion_Porcentaje(string List_codigos, int servicio)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Proceso_Verificacion_Porcentaje(List_codigos, servicio);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Negocio_Get_Alertas_lecturas(string fecha_asignacion , int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_Alertas_lecturas(fecha_asignacion, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        public string Capa_Negocio_Proceso_actualizar_lecturas(int id_lectura, string valor_lectura, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Proceso_Actualizar_lecturas(id_lectura, valor_lectura, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_Proceso_almacenar_lecturas( int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Proceso_Almacenar_lecturas( id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        public object Capa_Negocio_cambiarfoto(int id_usuario, int idfotoLectura, string nombrefotoLectura)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_cambiarfoto(id_usuario, idfotoLectura, nombrefotoLectura);
            }
            catch (Exception e)
            {
                throw e;
            }

        }



        public string Capa_Negocio_Proceso_almacenar_lecturas_vacias(int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Proceso_Almacenar_lecturas_vacias(id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> Capa_Negocio_ListandoLectura_Seguimiento(int id_local, int id_servicio, string fechainicial, string fechafinal, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_ListandoLectura_Seguimiento(id_local, id_servicio, fechainicial, fechafinal, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_ListandoFotosDescargar_seguimiento(int local, int servicio, string fecha,string option, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Negocio_ListandoFotosDescargar_seguimiento(local, servicio, fecha, option, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_Listando_Data_Descargar_seguimiento(int local, int servicio, string fecha, string option, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Negocio_Listando_Data_Descargar_seguimiento(local, servicio, fecha, option, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <returns></returns>
        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaSupervisor()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaSupervisor();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaSupervisor_Usuario(int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaSupervisor_usuario(id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaSupervisor_Usuario_validacion(int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaSupervisor_usuario_validacion(id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaOperarios()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaOperarios();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaUsuariosOperarios(int id_supervisor)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaUsuarioOperario(id_supervisor);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        

        public string Capa_Negocio_set_agregarOperario(string objOperarios, int id_supervisor, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_set_agregarOperario(objOperarios, id_supervisor, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_set_eliminarOperario(string objOperarios, int id_supervisor, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_set_eliminarOperario(objOperarios, id_supervisor, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
                

       public string Capa_Negocio_set_anulando_unidad_lectura(string Cod_UnidadLectura,  int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_set_anulando_unidad_lectura(Cod_UnidadLectura, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaAnio()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaAnio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_ListaMes()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_ListaMes();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_unidadlectura()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_UnidadLectura();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_unidadlectura_principal()
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_UnidadLectura_principal();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_set_Insert_Update_Configuracion(int id_Configuracion_UL, int dia_configuracion_ul, string cod_unidadlectura, int id_usuario_responsable, int id_tiposervicio, int estado, int usuario_creacion)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_set_Insert_Update_Configuracion(id_Configuracion_UL, dia_configuracion_ul, cod_unidadlectura, id_usuario_responsable, id_tiposervicio, estado, usuario_creacion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_set_Insert_Update_Unidad_lectura(string codigo, string nombre, string distrito, string estado, int usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_set_Insert_Update_Unidad_lectura(codigo, nombre, distrito, estado, usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Negocio_Get_Configuracion_ul(string dia_config, int id_supervisor, int id_servicio, int anio, int mes)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_Get_Configuracion_ul(dia_config, id_supervisor, id_servicio, anio, mes);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_set_generando_relectura(int servicio, string fecha_asignacion, int id_usuario)
        {
            try
            {
                Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
                return Objeto_Dato.Capa_Dato_set_generando_relectura(servicio, fecha_asignacion, id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_negocio_DescargarArchivoTexto_EnvioCliente_relectura(int local, int servicio, int estado, string fechaAsigna, int id_supervisor, int id_operario_supervisor)
        {
            Cls_Dato_AsignaOrdenTrabajo Objeto_Dato = new Cls_Dato_AsignaOrdenTrabajo();
            return Objeto_Dato.Capa_dato_DescargarArchivoTexto_EnvioCliente_Relectura(local, servicio, estado, fechaAsigna, id_supervisor, id_operario_supervisor);
        }


    }
}
