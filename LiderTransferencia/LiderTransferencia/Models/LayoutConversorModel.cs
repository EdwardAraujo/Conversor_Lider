using Move;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace LiderTransferencia.Models
{
    public class LayoutConversorModel
    {

        [Required]
        public int IdLayoutConversor { get; set; }

        public string Descricao { get; set; }


        public DataTable RetornaLayouts()
        {

            SqlHelper sql = new SqlHelper();

            DataTable t = new DataTable();

            t = sql.ExecuteQueryDataTable("select IdLayoutConversor,Descricao from LayoutConversor");

            return t;        
        
        }


    }
}
