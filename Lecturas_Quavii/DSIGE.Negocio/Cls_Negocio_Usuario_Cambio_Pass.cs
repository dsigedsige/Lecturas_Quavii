using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;


namespace DSIGE.Negocio
{
   public class Cls_Negocio_Usuario_Cambio_Pass
    {

       public List<Cls_entidad_Usuario_Cambio_Pass> Capa_Negocio_Verificando_Pass(int id_usuario, string contraseña)
        {
            try
            {
                Cls_Dato_Usuario_Cambio_Pass Objeto_Dato = new Cls_Dato_Usuario_Cambio_Pass();
                return Objeto_Dato.Capa_Dato_Verificando_Pass(id_usuario, contraseña);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


       public bool Capa_Negocio_Guardar_Informacion(Cls_entidad_Usuario_Cambio_Pass Entidad)
       {
           try
           {
               Cls_Dato_Usuario_Cambio_Pass Objeto_Dato = new Cls_Dato_Usuario_Cambio_Pass();
               return Objeto_Dato.Capa_Dato_Guardar_Informacion(Entidad);
           }
           catch (Exception e )
           {
            throw e;
           }
       }



    }
}
