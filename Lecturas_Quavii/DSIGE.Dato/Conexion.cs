using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Dato
{
    public class Conexion
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-03
        /// </summary>
        public Conexion() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-02
        /// </summary>
        /// <param name="__t"></param>
        /// <param name="__d"></param>
        /// <returns></returns>
        private int ExecuteBulkCopy(string __t, DataTable __d)
        {
            int filas = 0;

            try
            {
                filas = __d.Rows.Count;

                using (DbConnection dbconnection = DatabaseFactory.CreateDatabase().CreateConnection())
                {
                    dbconnection.Open();

                    if (filas > 0)
                    {
                        using (SqlBulkCopy bCopy = new SqlBulkCopy((SqlConnection)dbconnection))
                        {
                            bCopy.BulkCopyTimeout = 5000;
                            bCopy.DestinationTableName = __t;
                            bCopy.WriteToServer(__d);
                        }
                    }
                }

                return filas;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-02
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="__t"></param>
        /// <param name="__l"></param>
        /// <returns></returns>
        public int ExecuteBulkCopy<T>(string __t, IList<T> __l)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));

            DataTable table = new DataTable("Tabla");

            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (T item in __l)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }

                table.Rows.Add(row);
            }

            return ExecuteBulkCopy(__t, table);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-02
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="__t"></param>
        /// <param name="__l"></param>
        /// <returns></returns>
        public int ExecuteBulkCopy<T>(string __t, IEnumerable<T> __l)
        {
            PropertyInfo[] properties = typeof(T).GetProperties();

            DataTable table = new DataTable();

            foreach (var prop in properties)
            {
                table.Columns.Add(prop.Name, prop.PropertyType);
            }

            foreach (var item in __l)
            {
                DataRow row = table.NewRow();

                foreach (var prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item, null);
                }

                table.Rows.Add(row);
            }

            return ExecuteBulkCopy(__t, table); ;
        }
    }
}
