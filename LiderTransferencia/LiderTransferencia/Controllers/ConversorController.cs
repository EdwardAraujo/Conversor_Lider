using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using LiderTransferencia.Models;
using Microsoft.AspNetCore.Mvc;
using LiderTransferencia.Conversores;
using System.IO;

namespace LiderTransferencia.Controllers
{
    public class ConversorController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {

            ViewBag.ListaLayoutConversor = new LayoutConversorModel().RetornaLayouts();


            return View();
        }

      

        [HttpPost]
        public IActionResult Index(ConversorModel model)
        {

            switch (model.IdLayoutConversor) 
            {
                case 2:
                    LiderTransf conversor = new LiderTransf();
                    conversor.ConversorLayout(model);
                    break;
                default:
                    break;

            
            }

            string CaminhoArquivo = Path.Combine(@"C:\MOVEWEB\", "Arquivo" + Path.GetExtension(model.file.FileName));

            byte[] fileBytes = System.IO.File.ReadAllBytes(CaminhoArquivo.Replace("Arquivo", "V6LIDER"));

            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "Convertido_V6_" + model.file.FileName);

            //return RedirectToAction("Index");
        }

    }
}
