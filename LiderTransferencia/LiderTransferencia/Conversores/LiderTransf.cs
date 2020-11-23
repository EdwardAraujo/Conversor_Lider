using ClosedXML.Excel;
using LiderTransferencia.Models;
using Move;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace LiderTransferencia.Conversores
{
    public class LiderTransf
    {

        public void ConversorLayout(ConversorModel model) 
        {        

            if (model.file == null)
            {
                //TempData["ErrorModel"] = "Selecione um arquivo";
            }

            if (!Directory.Exists(@"C:\MOVEWEB\"))
            {
                //Criamos um com o nome folder
                Directory.CreateDirectory(@"C:\MOVEWEB");
            }

            string CaminhoArquivo = Path.Combine(@"C:\MOVEWEB\", "Arquivo" + Path.GetExtension(model.file.FileName));
            
            using (var fileStream = new FileStream(CaminhoArquivo, FileMode.Create, FileAccess.Write))
            {
                model.file.CopyTo(fileStream);

            }

            FileInfo V6REDE = new FileInfo(@"J:\Tecnologia_Operacional\SP\Publico\Sistemas\C#\Move\Layout_DOR_v6.xlsx");
            V6REDE.CopyTo(CaminhoArquivo.Replace("Arquivo", "V6LIDER"));

            var wbOrigem = new XLWorkbook(CaminhoArquivo);
            var wsOrigem = wbOrigem.Worksheet(1);

            var wbV6 = new XLWorkbook(CaminhoArquivo.Replace("Arquivo", "V6LIDER"));
            var wsV6 = wbV6.Worksheet(9);

            //Converte

            int linha = 7;
            int linhaV6 = 2;
            while (true)
            {
                //Valida se linha esta vazia e encerra
                if (string.IsNullOrEmpty(wsOrigem.Cell("A" + linha.ToString()).Value.ToString())) break;

                //Contrato, Matricula
                string Contrato = wsOrigem.Cell("A" + linha.ToString()).Value.ToString().Replace("-", "");                
                string MatriculaD = wsOrigem.Cell("E" + linha.ToString()).Value.ToString();

                DataTable tit = RetornaTitulares(Contrato, MatriculaD);

                // leitura do data table + select dos campos - Tabela Titular
                DataRow[] oDataRow = tit.Select("Ativo = 'SIM'");

                var vContrato = "";
                var vSubfatura = "";                
                var vMatricula = "";             
                var vCPF = "";      

                foreach (var item in oDataRow)
                {
                    vContrato = item[0].ToString();
                    vSubfatura = item[1].ToString();                    
                    vMatricula = item[3].ToString();                  
                    vCPF = item[7].ToString();
                }

                //TRATA CAMPOS DE RETORNO DO BANCO SQL VAZIOS
                String MsgNloc = "INFORMAÇÃO NÃO LOCALIZADA NO MOVE (MARCA OTICA:" + wsOrigem.Cell("E" + linha.ToString()).Value.ToString() + ")";
                if (vContrato == "")
                { vContrato = MsgNloc; }

                if (vSubfatura == "")
                { vSubfatura = MsgNloc; }               

                if (vMatricula == "")
                { vMatricula = MsgNloc; }   

                if (vCPF == "")
                { vCPF = MsgNloc; }

                //Operadora
                wsV6.Cell("A" + linhaV6.ToString()).Value = "Amil";
                wsV6.Cell("B" + linhaV6.ToString()).Value = vContrato;
                wsV6.Cell("C" + linhaV6.ToString()).Value = vSubfatura;
                wsV6.Cell("D" + linhaV6.ToString()).Value = wsOrigem.Cell("I" + linha.ToString()).Value;
                wsV6.Cell("E" + linhaV6.ToString()).Value = vMatricula;
                wsV6.Cell("G" + linhaV6.ToString()).Value = vCPF;
                wsV6.Cell("H" + linhaV6.ToString()).Value = wsOrigem.Cell("H" + linha.ToString()).Value; ;
               

                linha++;
                linhaV6++;

            }

            //Finaliza os arquivos
            wbV6.Save();
            wsOrigem = null;
            wbOrigem = null;
            wbV6 = null;
            wsV6 = null;


        }
      

        public DataTable RetornaTitulares(string Contrato, string MatriculaFunc)
        {
            SqlHelper sql = new SqlHelper();
            DataTable tit = new DataTable();
            string Query = $"SELECT CONTRATO, SUBFATURA, CARTAO, MATRICULA_FUNC, NOME_TITULAR, CPF_TITULAR, NOME, CPF, ATIVO FROM Beneficiario WITH (NOLOCK) WHERE Contrato like '%" + Contrato + "%' AND CARTAO = '" + MatriculaFunc + "'";

            tit = sql.ExecuteQueryDataTable(Query);

            return tit;

        }




    }
}
