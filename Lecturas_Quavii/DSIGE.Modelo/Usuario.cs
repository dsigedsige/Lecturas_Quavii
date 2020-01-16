using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Usuario
    {
        public int usu_id { get; set; }

        public string usu_nrodoc { get; set; }

        public string usu_apellidos { get; set; }

        public string usu_nom { get; set; }

        public string usu_email { get; set; }

        public string usu_usuario { get; set; }

        public string usu_clave { get; set; }

        public string usu_tipo { get; set; }

        public string usu_estado { get; set; }

        public string usu_cargo { get; set; }

        public string usu_check { get; set; }

        public string usu_idperfil { get; set; }

        public int usu_usuCrea { get; set; }

        public int usu_usuEdicion { get; set; }

        public string permisos_opciones { get; set; }

        public string usu_perfil { get; set; }

    }

    /// <summary>
    /// Autor: rcontreras
    /// Fecha: 27-09-2015
    /// </summary>
    public class Request_Usuario_Recupera_Clave
    {
        [JsonProperty("_a")]
        public string usu_email { get; set; }
    }

    public class Usuario_Servicio
    {
        public int idOperario_Servicio { get; set; }

        public String nombre_tiposervicio { get; set; }

        public int idOperario { get; set; }

        public int id_TipoServicio { get; set; }

        public int estado { get; set; }

        public DateTime fechaCreacion { get; set; }

        public String descripcion { get; set; }
        public String color { get; set; }
        public String icono { get; set; }


    }

}
