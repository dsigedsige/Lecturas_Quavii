using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Data.SqlClient;

namespace DSIGE.Dato
{
    public class DUbicacion_Lecturas
    {

        public DUbicacion_Lecturas() {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<UbicacionOperario> DListaLecturasOperariosGPS(int idOpe, string fechaAsig, string suministro, string medidor, int idEmp,string lista)
        {
            try
            {
                List<UbicacionOperario> lOb = new List<UbicacionOperario>();

                string[] Servicios = lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                int total = Servicios.Length;

                if (total == 1) //UN SERVICIO
                {
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_UBICA_LECTURAS", idOpe, fechaAsig, suministro, medidor, idEmp, Servicios[0]))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new UbicacionOperario()
                                        {
                                            idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                            id_ope = Convert.ToInt32(iDr["id_Operario_Lectura"]),
                                            ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                            suministro = Convert.ToString(iDr["suministro_lectura"]),
                                            medidor = Convert.ToString(iDr["medidor_lectura"]),
                                            direcLectura = Convert.ToString(iDr["direccion_lectura"]),
                                            nom_cliente = Convert.ToString(iDr["nombreCliente_lectura"]),
                                            lectura = Convert.ToString(iDr["lectura"]),
                                            foto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                            lec_estado = Convert.ToInt32(iDr["estado"]),
                                            latitud = Convert.ToString(iDr["latitud_lectura"]),
                                            longitud = Convert.ToString(iDr["longitud_lectura"]),
                                            fasig = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                            ope_tipo = Convert.ToString(iDr["tipoUsuario"]),
                                            fmovil = Convert.ToString(iDr["fechaLecturaMovil_lectura"])
                                        }
                                    );
                            }
                        }
                    }
                }
                else
                { // TODOS LOS SERVICIOS
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_UBICA_LECTURA_TODOS_SERVICIOS", idOpe, fechaAsig, suministro, medidor, idEmp))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new UbicacionOperario()
                                        {
                                            idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                            id_ope = Convert.ToInt32(iDr["id_Operario_Lectura"]),
                                            ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                            suministro = Convert.ToString(iDr["suministro_lectura"]),
                                            medidor = Convert.ToString(iDr["medidor_lectura"]),
                                            direcLectura = Convert.ToString(iDr["direccion_lectura"]),
                                            nom_cliente = Convert.ToString(iDr["nombreCliente_lectura"]),
                                            lectura = Convert.ToString(iDr["lectura"]),
                                            foto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                            lec_estado = Convert.ToInt32(iDr["estado"]),
                                            latitud = Convert.ToString(iDr["latitud_lectura"]),
                                            longitud = Convert.ToString(iDr["longitud_lectura"]),
                                            fasig = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                            ope_tipo = Convert.ToString(iDr["tipoUsuario"]),
                                            fmovil = Convert.ToString(iDr["fechaLecturaMovil_lectura"])
                                        }
                                    );
                            }
                        }
                    } 
                }
                




                return lOb;
            }
            catch (Exception e)
            {
                throw e;
 
            }
        }


        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

        public List<Ubicacion_Lectura> Capa_Dato_Get_ListaOperarios()
        {
            try
            {
                List<Ubicacion_Lectura> Listlocal = new List<Ubicacion_Lectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_OPERARIO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Ubicacion_Lectura obj_entidad = new Ubicacion_Lectura();

                                obj_entidad.id_Operario = Convert.ToInt32(Fila["id_Operario"].ToString());
                                obj_entidad.desc_operario = Fila["desc_operario"].ToString();
                                Listlocal.Add(obj_entidad);
                            }
                        }
                    }
                }

                return Listlocal;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Ubicacion_Lectura> Capa_Dato_Get_ListaServicio()
        {
            try
            {
                List<Ubicacion_Lectura> ListServicio = new List<Ubicacion_Lectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Ubicacion_Lectura obj_entidad = new Ubicacion_Lectura();

                                obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                ListServicio.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListServicio;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<Ubicacion_Lectura> Capa_Dato_Get_ListaServicio_Efectividad()
        {
            try
            {
                List<Ubicacion_Lectura> ListServicio = new List<Ubicacion_Lectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO_REPORTE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Ubicacion_Lectura obj_entidad = new Ubicacion_Lectura();

                                obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                ListServicio.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListServicio;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<UbicacionOperario> DSeguimiento_Lectura_II(int Local , string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            List<UbicacionOperario> lOb = new List<UbicacionOperario>();
            try
            {

                if (servicio > 0) //UN SERVICIO
                {
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_UBICA_LECTURAS", operario, pFechaAsiga, suministro, medidor, 1, servicio, Local))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new UbicacionOperario()
                                        {
                                            idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                            id_ope = Convert.ToInt32(iDr["id_Operario_Lectura"]),
                                            ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                            suministro = Convert.ToString(iDr["suministro_lectura"]),
                                            medidor = Convert.ToString(iDr["medidor_lectura"]),
                                            direcLectura = Convert.ToString(iDr["direccion_lectura"]),
                                            nom_cliente = Convert.ToString(iDr["nombreCliente_lectura"]),
                                            lectura = Convert.ToString(iDr["lectura"]),
                                            foto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                            lec_estado = Convert.ToInt32(iDr["estado"]),
                                            latitud = Convert.ToString(iDr["latitud_lectura"]),
                                            longitud = Convert.ToString(iDr["longitud_lectura"]),
                                            fasig = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                            ope_tipo = Convert.ToString(iDr["tipoUsuario"]),
                                            fmovil = Convert.ToString(iDr["fechaLecturaMovil_lectura"])
                                        }
                                    );
                            }
                        }
                    }
                }
                else
                { // TODOS LOS SERVICIOS
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_UBICA_LECTURA_TODOS_SERVICIOS", operario, pFechaAsiga, suministro, medidor, 1))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new UbicacionOperario()
                                        {
                                            idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                            id_ope = Convert.ToInt32(iDr["id_Operario_Lectura"]),
                                            ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                            suministro = Convert.ToString(iDr["suministro_lectura"]),
                                            medidor = Convert.ToString(iDr["medidor_lectura"]),
                                            direcLectura = Convert.ToString(iDr["direccion_lectura"]),
                                            nom_cliente = Convert.ToString(iDr["nombreCliente_lectura"]),
                                            lectura = Convert.ToString(iDr["lectura"]),
                                            foto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                            lec_estado = Convert.ToInt32(iDr["estado"]),
                                            latitud = Convert.ToString(iDr["latitud_lectura"]),
                                            longitud = Convert.ToString(iDr["longitud_lectura"]),
                                            fasig = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                            ope_tipo = Convert.ToString(iDr["tipoUsuario"]),
                                            fmovil = Convert.ToString(iDr["fechaLecturaMovil_lectura"])
                                        }
                                    );
                            }
                        }
                    }
                }
                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<UbicacionOperario> DSeguimiento_Lectura_III(string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            List<UbicacionOperario> lOb = new List<UbicacionOperario>();
            try
            {
                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("PROC_SEGUIMIENTO_OPERARIO_DATO_SUMINISTRO", operario, pFechaAsiga, suministro, medidor, 1, servicio))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new UbicacionOperario()
                                    {
                                        idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                        id_ope = Convert.ToInt32(iDr["id_Operario_Lectura"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        suministro = Convert.ToString(iDr["suministro_lectura"]),
                                        medidor = Convert.ToString(iDr["medidor_lectura"]),
                                        direcLectura = Convert.ToString(iDr["direccion_lectura"]),
                                        nom_cliente = Convert.ToString(iDr["nombreCliente_lectura"]),
                                        lectura = Convert.ToString(iDr["lectura"]),
                                        foto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                        lec_estado = Convert.ToInt32(iDr["estado"]),
                                        latitud = Convert.ToString(iDr["latitud_lectura"]),
                                        longitud = Convert.ToString(iDr["longitud_lectura"]),
                                        fasig = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                        ope_tipo = Convert.ToString(iDr["tipoUsuario"]),
                                        fmovil = Convert.ToString(iDr["fechaLecturaMovil_lectura"])
 
                                    }
                                );
                        }
                    }
                }
                
                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<UbicacionOperario> DSeguimiento_Lectura_Resumen(string FechaAsiga, int servicio, int operario)
        {
            List<UbicacionOperario> lOb = new List<UbicacionOperario>();
            try
            {
                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_UBICA_LECTURAS_RESUMEN", FechaAsiga, servicio, operario))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new UbicacionOperario()
                                        {
                                            idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                            id_ope = Convert.ToInt32(iDr["id_Operario_Lectura"]),
                                            ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                            suministro = Convert.ToString(iDr["suministro_lectura"]),
                                            medidor = Convert.ToString(iDr["medidor_lectura"]),
                                            direcLectura = Convert.ToString(iDr["direccion_lectura"]),
                                            nom_cliente = Convert.ToString(iDr["nombreCliente_lectura"]),
                                            lectura = Convert.ToString(iDr["lectura"]),
                                            foto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                            lec_estado = Convert.ToInt32(iDr["estado"]),
                                            latitud = Convert.ToString(iDr["latitud_lectura"]),
                                            longitud = Convert.ToString(iDr["longitud_lectura"]),
                                            fasig = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                            ope_tipo = Convert.ToString(iDr["tipoUsuario"]),
                                            fmovil = Convert.ToString(iDr["fechaLecturaMovil_lectura"])
                                        }
                                    );
                            }
                        }
                    }
                        
                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Ubicacion_Lectura> Capa_Dato_Get_UbicacionLecturas(int Local,  string fechaAsigna, int servicio, int operario)
        {
            try
            {
                List<Ubicacion_Lectura> Obj_Lista = new List<Ubicacion_Lectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_REPORTE_UBICA_LECTURAS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Local", SqlDbType.Int).Value = Local;
                        cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@operario", SqlDbType.Int).Value = operario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Ubicacion_Lectura obj_entidad = new Ubicacion_Lectura();

                                obj_entidad.idLectura = Convert.ToInt32(Fila["id_Lectura"]);
                                obj_entidad.id_ope = Convert.ToInt32(Fila["id_Operario_Lectura"]);
                                obj_entidad.ope_nombre = Convert.ToString(Fila["ope_nombre"]);
                                obj_entidad.suministro = Convert.ToString(Fila["suministro_lectura"]);
                                obj_entidad.medidor = Convert.ToString(Fila["medidor_lectura"]);
                                obj_entidad.direcLectura = Convert.ToString(Fila["direccion_lectura"]);
                                obj_entidad.nom_cliente = Convert.ToString(Fila["nombreCliente_lectura"]);
                                obj_entidad.lectura = Convert.ToString(Fila["lectura"]);
                                obj_entidad.foto_lectura = Convert.ToString(Fila["tieneFoto_lectura"]);
                                obj_entidad.lec_estado = Convert.ToInt32(Fila["estado"]);
                                obj_entidad.latitud = Convert.ToString(Fila["latitud_lectura"]);
                                obj_entidad.longitud = Convert.ToString(Fila["longitud_lectura"]);
                                obj_entidad.fasig = Convert.ToString(Fila["fechaAsignacion_Lectura"]);
                                obj_entidad.ope_tipo = Convert.ToString(Fila["tipoUsuario"]);
                                obj_entidad.fmovil = Convert.ToString(Fila["fechaLecturaMovil_lectura"]);
                                Obj_Lista.Add(obj_entidad);
                            }
                        }
                    }
                }

                return Obj_Lista;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
