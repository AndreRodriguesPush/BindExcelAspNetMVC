using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML;
using ClosedXML.Excel;

namespace ExtracaoExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {

            return View();
        }

        public ActionResult UploadFile(HttpPostedFileBase Anexo)
        {
            try
            {
                string filePath = string.Empty;

                if (Anexo != null && Anexo.ContentLength > 0)
                {
                    //if (Anexo.FileName.EndsWith(".csv"))
                    //{
                    string path = Server.MapPath("~/ArquivoUpload/");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    //Salva arquivo em diretório dentro da aplicação
                    filePath = path + Path.GetFileName(Anexo.FileName);
                    string extension = Path.GetExtension(Anexo.FileName);
                    Anexo.SaveAs(filePath);

                    //Lendo o conteudo do arquivo csv
                    var arquivoExcel = System.IO.File.ReadAllText(filePath);

                    var localArquivo = new XLWorkbook(filePath);
                    var planilha = localArquivo.Worksheet(1);

                    var linhaMesReferencia = 3;
                    var MesReferencia = planilha.Cell("A" + linhaMesReferencia.ToString()).Value.ToString();

                    var linhaCampus = 4;
                    var Campus = planilha.Cell("A" + linhaCampus.ToString()).Value.ToString();


                    List<ViewModelBolsista> vwBolsista = new List<ViewModelBolsista>();

                    var linhaNome = 6;
                    while (true)
                    {
                        var NomeAluno = planilha.Cell("B" + linhaNome.ToString()).Value.ToString();
                        linhaNome++;

                    }



                    //}
                }


            }
            catch (Exception)
            {

            }

            return View();
        }

    }
}