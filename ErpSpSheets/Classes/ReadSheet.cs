﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ErpSpSheets
{
    [Guid("69EE8D62-75B7-4633-9B4B-9DB7F16D5937"), ClassInterface(ClassInterfaceType.None), ComVisible(true)]
    public class ReadSheet : IReadSheet
    {

        private ExcelWorksheet ws;
        private ExcelPackage package;
        private int lastRow;
        private int lastCol;
        private string retType;

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

        public string ReadText(int ln, int col)
        {
            string retText = "";
            bool podeLer = true;
            
            if (ws != null)
            {
                lastRow = ws.Dimension.End.Row;
                lastCol = ws.Dimension.End.Column;

                if (ln > lastRow || col > lastCol || ws.Cells[ln, col].Value == null)
                    podeLer = false;

                if (podeLer && !string.IsNullOrEmpty(ws.Cells[ln, col].Value.ToString()))
                {
                    retText = ws.Cells[ln, col].Value.ToString();
                }
            }
            
            return retText;
        }

        public int ReadNumber(int ln, int col)
        {
            int retNumber = 0;
            bool podeLer = true;

            if (ws != null)
            {
                lastRow = ws.Dimension.End.Row;
                lastCol = ws.Dimension.End.Column;

                if (ln > lastRow || col > lastCol || ws.Cells[ln, col].Value == null)
                    podeLer = false;

                if (podeLer && !string.IsNullOrEmpty(ws.Cells[ln, col].Value.ToString()))
                {
                    try
                    {
                        retNumber = Convert.ToInt32(ws.Cells[ln, col].Value.ToString());
                    }
                    catch (Exception)
                    {
                        retNumber = 0;
                    }
                }
            }

            return retNumber;
        }

        public decimal ReadReal(int ln, int col)
        {
            decimal retReal = 0;
            bool podeLer = true;

            if (ws != null)
            {                
                try
                {
                    lastRow = ws.Dimension.End.Row;
                    lastCol = ws.Dimension.End.Column;

                    if (ln > lastRow || col > lastCol || ws.Cells[ln, col].Value == null)
                        podeLer = false;

                    if (podeLer)
                    {
                        retReal = Convert.ToDecimal(ws.Cells[ln, col].Value.ToString());
                    }
                }
                catch (Exception)
                {
                    retReal = 0;
                }
            }

            return retReal;
        }

        public DateTime ReadDate(int ln, int col)
        {
            DateTime retDate = new DateTime();
            bool podeLer = true;

            if (ws != null)
            {
                try
                {
                    lastRow = ws.Dimension.End.Row;
                    lastCol = ws.Dimension.End.Column;

                    if (ln > lastRow || col > lastCol || ws.Cells[ln, col].Value == null)
                        podeLer = false;

                    if (podeLer)
                    {
                        retDate = Convert.ToDateTime(ws.Cells[ln, col].Value.ToString());
                    }
                }
                catch (Exception)
                {
                    retDate = Convert.ToDateTime("1753-01-01");
                }
            }

            return retDate;
        }

        public string ReadType(int ln, int col)
        {
            string retType = "";
            bool podeLer = true;

            if (ws != null)
            {
                lastRow = ws.Dimension.End.Row;
                lastCol = ws.Dimension.End.Column;

                if (ln > lastRow || col > lastCol || ws.Cells[ln, col].Value == null)
                    podeLer = false;

                if (podeLer && !string.IsNullOrEmpty(ws.Cells[ln, col].Value.ToString()))
                {
                    var tipo = ws.Cells[ln, col].Value.GetType();
                    retType = tipo.ToString();
                }
            }

            return retType;
        }

        public string ClosePlanilha ()
        {
            string result = "";

            if (ws != null)
            {
                try
                {
                    package.Dispose();
                }
                catch (Exception ex)
                {
                    result = "Não Foi Possivel Fechar a Planilha! \n" + ex.Message;
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
