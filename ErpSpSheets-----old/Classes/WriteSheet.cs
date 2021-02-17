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

        public string OpenPlanilha(string FileName, string WorkSheet, string tipo)
         {
            string result = "";

            try
            {
                if (!string.IsNullOrEmpty(WorkSheet) && WorkSheet.Length <= 32)
                {
                    //Define Uso Não Comercial
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    result = VerArquivo(FileName);
                    if (string.IsNullOrEmpty(result)) //sem erros
                    {
                        package = new ExcelPackage(new FileInfo(FileName));
                        ws = package.Workbook.Worksheets.Add(WorkSheet); //Cria a Planilha                    
                    }
                }
                else
                {
                    result = "[ERRO] Favor Informar o Nome da Planilha com no Maximo 32 Caracteres";
                }
            }
            catch (Exception ex)
            {
                result = "[ERRO] Não Foi Possivel Abrir a Planilha! \n" + ex.Message;
            }

            return result;
        }

        public string CellValue(int ln, int col, string value)
        {
            string retorno = "";
            try
            {
                ws.Cells[ln, col].Value = value;

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string CellFormula(int ln, int col, string value)
        {
            string retorno = "";
            try
            {
                ws.Cells[ln, col].Formula = value;

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string CellHAlign(int ln, int col, string value)
        {
            string retorno = "";

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

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string CellVAlign(int ln, int col, string value)
        {
            string retorno = "";

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

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string CellStyle(int ln, int col, string style, string value)
        {
            string retorno = "";

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
                }

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string RangeValue(string range, string value)
        {
            string retorno = "";

            try
            {
                ws.Cells[range].Value = value;

                return retorno;
            }

            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string RangeFormula(string range, string value)
        {
            string retorno = "";
            try
            {
                ws.Cells[range].Formula = value;

                return retorno;
            }

            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string RangeHAlign(string range, string value)
        {
            string retorno = "";
            try
            {
                switch (value)
                {
                    case "C": //Centro
                        ws.Cells[range].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        break;
                    case "E": //Esquerdo
                        ws.Cells[range].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        break;
                    case "D": //Direito
                        ws.Cells[range].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        break;
                    case "J": //Justificado
                        ws.Cells[range].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Justify;
                        break;
                    default:  //Esquerdo
                        ws.Cells[range].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        break;
                }

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string RangeVAlign(string range, string value)
        {
            string retorno = "";

            try
            {
                switch (value)
                {
                    case "C": //Centro
                        ws.Cells[range].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        break;
                    case "E": //Em Baixo
                        ws.Cells[range].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                        break;
                    case "D": //Direito
                        ws.Cells[range].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        break;
                    default:  //Esquerdo
                        ws.Cells[range].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                        break;
                }

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string RangeImage(string patchImg, int colImg, int rowImg, int posTop, int posLeft, int imgWidth, int imgHeight)
        {
            string range = "";
            string retorno = "";

            try
            {
                if (string.IsNullOrEmpty(patchImg))
                    retorno = "[ERRO] Favor informe o caminho da imagem! Parâmetro 1 PatchImg.";

                if (colImg == 0)
                    retorno = "[ERRO] Favor informe qual a Coluna que a Imagem será inserida! Parâmetro 2 ColImg.";

                if (rowImg == 0)
                    retorno = "[ERRO] Favor informe qual a Linha que a Imagem será inserida! Parâmetro 3 RowImg.";

                if (string.IsNullOrEmpty(retorno))
                {
                    range = ColunaTxt(colImg) + ":" + ColunaTxt(rowImg);
                    Image LogoEmp = Image.FromFile(patchImg);

                    if (imgWidth == 0)
                        imgWidth = LogoEmp.Width;

                    if (imgHeight == 0)
                        imgHeight = LogoEmp.Height;


                    ws.Cells[range].Merge = true;
                    OfficeOpenXml.Drawing.ExcelPicture picture = ws.Drawings.AddPicture("0", LogoEmp);
                    picture.From.Column = colImg;
                    picture.From.Row = rowImg;
                    picture.SetPosition(posTop, posLeft);
                    picture.SetSize(imgWidth, imgHeight);
                }

                return retorno;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }
        }

        public string RangeMerge(string range, string value)
        {
            string retorno = "";
            try
            {
                ws.Cells[range].Merge = (value == "true") ? true : false;
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }

            return retorno;
        }

        public string RangeStyle(string range, string style, string value)
        {
            string retorno = "";
            string info = "";

            info = "Método: RangeStyle \n" +
                "Parâmetros: \n" +
                "   string range: Range de Celulas que serão Afetadas. Ex: 'A1', 'A1:A5', 'A1:C5', etc\n" +
                "   string style: Stilo que será Afetado, Opções Disponíveis abaixo:\n" +
                "       NumberFormat - Formata o Tipo de valor da Celula. Ex. '@'(texto), '0'(Número), 'dd/MM/yyyy'(Data), para mais exemplos veja as Configurações de Celula do Próprio Excel\n" +
                "       Font.Bold - Deixa a Fonte do Range Negrito ou não. Passar value 'true' ou 'false'\n" +
                "       Font.Size - Altera o Tamanho da Fonte do Range. Passar um valor inteiro no value para o tamanho\n" +
                "       Font.Name - Altera a Fonte utilizada no Range. Passar o Literal nome da Fonete em value\n" +
                "   string value: Valor a ser atribuido ao Estilo informado.";

            if (value == "?") //Retorna a Documentação
                return info;

            try
            {
                switch (style)
                {
                    case "NumberFormat":
                        ws.Cells[range].Style.Numberformat.Format = value;
                        break;
                    case "Font.Bold":
                        ws.Cells[range].Style.Font.Bold = (value == "true") ? true : false;
                        break;
                    case "Font.Size":
                        ws.Cells[range].Style.Font.Size = (!string.IsNullOrEmpty(value)) ? Convert.ToInt32(value) : 12;
                        break;
                    case "Font.Name":
                        ws.Cells[range].Style.Font.Name = value;
                        break;
                    default:
                        retorno = "[ERRO] Stilo Não Implementado ou Não Existe.\n\n"+info;
                        break;
                }
            }
            catch (Exception ex)
            {
                retorno = "[ERRO] " + ex.Message;
                return retorno;
            }

            return retorno;
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
            if (ArquivoEmUso(ArqExcel))
                result = "[ERRO] Parece que a Planilha já esta Aberta! É necessário fechá-la antes de iniciar o processamento!";
            else
                if (Exfile.Exists)
                    Exfile.Delete();
          
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

        public string ColunaTxt(int col)
        {
            string[] ColunaExcel = new string[100];

            ColunaExcel[1] = "A";
            ColunaExcel[2] = "B";
            ColunaExcel[3] = "C";
            ColunaExcel[4] = "D";
            ColunaExcel[5] = "E";
            ColunaExcel[6] = "F";
            ColunaExcel[7] = "G";
            ColunaExcel[8] = "H";
            ColunaExcel[9] = "I";
            ColunaExcel[10] = "J";
            ColunaExcel[11] = "K";
            ColunaExcel[12] = "L";
            ColunaExcel[13] = "M";
            ColunaExcel[14] = "N";
            ColunaExcel[15] = "O";
            ColunaExcel[16] = "P";
            ColunaExcel[17] = "Q";
            ColunaExcel[18] = "R";
            ColunaExcel[19] = "S";
            ColunaExcel[20] = "T";
            ColunaExcel[21] = "U";
            ColunaExcel[22] = "V";
            ColunaExcel[23] = "W";
            ColunaExcel[24] = "X";
            ColunaExcel[25] = "Y";
            ColunaExcel[26] = "Z";
            ColunaExcel[27] = "AA";
            ColunaExcel[28] = "AB";
            ColunaExcel[29] = "AC";
            ColunaExcel[30] = "AD";
            ColunaExcel[31] = "AE";
            ColunaExcel[32] = "AF";
            ColunaExcel[33] = "AG";
            ColunaExcel[34] = "AH";
            ColunaExcel[35] = "AI";
            ColunaExcel[36] = "AJ";
            ColunaExcel[37] = "AK";
            ColunaExcel[38] = "AL";
            ColunaExcel[39] = "AM";
            ColunaExcel[40] = "AN";
            ColunaExcel[41] = "AO";
            ColunaExcel[42] = "AP";
            ColunaExcel[43] = "AQ";
            ColunaExcel[44] = "AR";
            ColunaExcel[45] = "AS";
            ColunaExcel[46] = "AT";
            ColunaExcel[47] = "AU";
            ColunaExcel[48] = "AV";
            ColunaExcel[49] = "AW";
            ColunaExcel[50] = "AX";
            ColunaExcel[51] = "AY";
            ColunaExcel[52] = "AZ";
            ColunaExcel[53] = "BA";
            ColunaExcel[54] = "BB";
            ColunaExcel[55] = "BC";
            ColunaExcel[56] = "BD";
            ColunaExcel[57] = "BE";
            ColunaExcel[58] = "BF";
            ColunaExcel[59] = "BG";
            ColunaExcel[60] = "BH";
            ColunaExcel[61] = "BI";
            ColunaExcel[62] = "BJ";
            ColunaExcel[63] = "BK";
            ColunaExcel[64] = "BL";
            ColunaExcel[65] = "BM";
            ColunaExcel[66] = "BN";
            ColunaExcel[67] = "BO";
            ColunaExcel[68] = "BP";
            ColunaExcel[69] = "BQ";
            ColunaExcel[70] = "BR";
            ColunaExcel[71] = "BS";
            ColunaExcel[72] = "BT";
            ColunaExcel[73] = "BU";
            ColunaExcel[74] = "BV";
            ColunaExcel[75] = "BW";
            ColunaExcel[76] = "BX";
            ColunaExcel[77] = "BY";
            ColunaExcel[78] = "BZ";

            return ColunaExcel[col];
        }                

    }
}
