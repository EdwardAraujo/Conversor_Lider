using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LiderTransferencia.Models
{
    public class ConversorModel
    {

        public int IdLayoutConversor { get; set; }

        public IFormFile file { get; set; }


    }
}
