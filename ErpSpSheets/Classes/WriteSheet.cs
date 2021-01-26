using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;

namespace ErpSpSheets
{
    [Guid("A9CE2D97-004D-4C40-8C06-8E5214CD57A7"), ClassInterface(ClassInterfaceType.None), ComVisible(true)]
    public class WriteSheet : IWriteSheet
    {
        private ExcelWorksheet ws;
        private ExcelPackage package;

        public string OpenPlanilha(string FileName, string WorkSheet)
        {
            string result = "";

            try
            {
                //Define Uso Não Comercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                result = VerArquivo(FileName);
                if (string.IsNullOrEmpty(result)) //sem erros
                {
                    package = new ExcelPackage(new FileInfo(FileName));

                    if (string.IsNullOrEmpty(WorkSheet))
                    {
                        ws = package.Workbook.Worksheets.First(); //Recupera a planilha para Leitura.
                    }
                    else
                    {
                        ws = package.Workbook.Worksheets[WorkSheet]; //Recupera a planilha para edição.
                    }
                }
            }
            catch (Exception ex)
            {
                result = "Não Foi Possivel Abrir a Planilha! \n" + ex.Message;
            }

            return result;
        }

        public void CellValue(int ln, int col, string value)
        {
            try
            {
                ws.Cells[ln, col].Value = value;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CellFormula(int ln, int col, string value)
        {
            try
            {
                ws.Cells[ln, col].Formula = value;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CellHAlign(int ln, int col, string value)
        {
            try
            {
                switch(value)
                {
                    case "C": //Centro
                        ws.Cells[ln, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        break;
                    case "E": //Esquerdo
                        ws.Cells[ln, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        break;
                    case "D": //Direito
                        ws.Cells[ln, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        break;
                    case "J": //Justificado
                        ws.Cells[ln, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Justify;
                        break;
                    default:  //Esquerdo
                        ws.Cells[ln, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CellVAlign(int ln, int col, string value)
        {
            try
            {
                switch (value)
                {
                    case "C": //Centro
                        ws.Cells[ln, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        break;
                    case "E": //Em Baixo
                        ws.Cells[ln, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                        break;
                    case "D": //Direito
                        ws.Cells[ln, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        break;
                    default:  //Esquerdo
                        ws.Cells[ln, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CellStyle(int ln, int col, string style, string value)
        {
            try
            {
                switch(style)
                {
                    case "NumberFormat":
                        ws.Cells[ln, col].Style.Numberformat.Format = value;
                        break;
                    case "Font.Bold":
                        ws.Cells[ln, col].Style.Font.Bold = (value == "true") ? true : false;
                        break;
                    case "Font.Size":
                        ws.Cells[ln, col].Style.Font.Size = (!string.IsNullOrEmpty(value)) ? Convert.ToInt32(value) : 12;
                        break;
                    case "Font.Name":
                        ws.Cells[ln, col].Style.Font.Name = value;
                        break;
                    case "BackCollor":
                        ws.Cells[ln, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[ln, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb();
                        break;

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void RangeValue(string range, string value)
        {
            throw new NotImplementedException();
        }

        public void RangeFormula(string range, string value)
        {
            throw new NotImplementedException();
        }

        public void RangeHAlign(string range, string value)
        {
            throw new NotImplementedException();
        }

        public void RangeImage(string patchImg, int colImg, int rowImg, int posTop, int posLeft, int imgWidth, int imgHeight)
        {
            throw new NotImplementedException();
        }

        public void RangeMerge(string range, string value)
        {
            throw new NotImplementedException();
        }

        public void RangeStyle(string range, string style, string value)
        {
            throw new NotImplementedException();
        }

        public void RangeVAlign(string range, string value)
        {
            throw new NotImplementedException();
        }

        public string SavePlanilha()
        {
            string result = "";

            if (ws != null)
            {
                try
                {
                    package.Save();
                }
                catch (Exception ex)
                {
                    result = "Não Foi Possivel Salvar a Planilha! \n" + ex.Message;
                }
            }

            return result;
        }

        private string VerArquivo(string ArqExcel)
        {
            string result = "";

            FileInfo Exfile = new FileInfo(ArqExcel);
            if (Exfile.Exists)
            {
                if (ArquivoEmUso(ArqExcel))
                    result = "[ERRO]Parece que a Planilha já esta Aberta! É necessário fechá-la antes de iniciar o processamento!";
            }
            else
            {
                result = "[ERRO]O arquivo: " + ArqExcel.Trim() + ", não foi encontrado! Impossivel iniciar o processamento.";
            }

            return result;
        }
        
        private bool ArquivoEmUso(string caminhoArquivo)
        {
            try
            {
                FileStream fs = File.OpenWrite(caminhoArquivo);
                fs.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }

        }
    }
}
