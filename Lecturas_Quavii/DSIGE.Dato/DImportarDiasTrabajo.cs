using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Practices.EnterpriseLibrary.Data;
using DSIGE.Modelo;
using System.Data.SqlClient;
using System.Data;
using System.Data.OleDb;

namespace DSIGE.Dato
{
  public  class DImportarDiasTrabajo
    {
        public DImportarDiasTrabajo()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        OleDbConnection cn;

        private OleDbConnection ConectarExcel(string nomExcel)
        {
            cn = new OleDbConnection();
            try
            {                
                cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + nomExcel + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'";
                cn.Open();
                return cn;
            }
            catch (Exception)
            {
                cn.Close();
                throw;
            }
        }

        public List<Local> getLocales()
        {
            try
            {
                List<Local> cargo = new List<Local>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getLocales"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            cargo.Add(
                                    new Local()
                                    {
                                        loc_id = Convert.ToInt32(iDr["id_Local"]),
                                        loc_nombre = Convert.ToString(iDr["nombre_local"]),
                                   }
                                );
                        }
                    }
                }
                return cargo;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ImportarDiasTrabajo> getSector_DiasTrabajo( int idLocal, string fechaIni, string fechaFin)
        {
            try
            {
                List<ImportarDiasTrabajo> cargo = new List<ImportarDiasTrabajo>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_getSector_DiasTrabajo",idLocal,fechaIni,fechaFin))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            cargo.Add(
                                    new ImportarDiasTrabajo()
                                    {
                                        id_Sector = Convert.ToInt32(iDr["id_Sector"]),
                                        Sector = Convert.ToString(iDr["Sector"]),
                                        fecha = Convert.ToString(iDr["fecha"]),
                                        estado = Convert.ToString(iDr["estado"]),
                                        id_Local = Convert.ToInt32(iDr["id_Local"])                                   
                                    }
                                );
                        }
                    }
                }
                return cargo;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool ActualizarSector_DiasTrabajo(ImportarDiasTrabajo objDiasTrabajo)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_ActualizarSector_DiasTrabajo",
                    objDiasTrabajo.Sector,
                    objDiasTrabajo.fecha,
                    objDiasTrabajo.estado,
                    objDiasTrabajo.usuario_edicion,
                    objDiasTrabajo.id_Sector);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public DataTable ListaExcel(string fileLocation)
        {
            DataTable dt = new DataTable();
            try
            {
                string sql = "SELECT *FROM [Importar$]";
           
                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();             
            }
            catch (Exception)
            {

                cn.Close();
            }
            return dt;
           
        }


        public bool RegistrarSector_DiasTrabajo(List<ImportarDiasTrabajo> ListDiasTrabajo)
        {
            try
            {
                foreach (ImportarDiasTrabajo item in ListDiasTrabajo)
                {
                  DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_RegistrarSector_DiasTrabajo",
                  item.Sector,
                  item.fecha,           
                  item.usuario_creacion,
                  item.id_Local);
                }             
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



    






    }
}
