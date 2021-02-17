using System;
using System.Runtime.InteropServices;

namespace ErpSpSheets
{
    [Guid("597E3623-207E-47FF-99E8-81F1F1560101"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), ComVisible(true)]
    public interface IWriteSheet
    {
        //string OpenPlanilha(string FileName, string WorkSheet);
        //string SavePlanilha();

        //string CellValue(int ln, int col, string value);
        //string CellFormula(int ln, int col, string value);
        //string CellHAlign(int ln, int col, string value);
        //string CellVAlign(int ln, int col, string value);
        //string CellStyle(int ln, int col, string style, string value);

        //string RangeValue(string range, string value);
        //string RangeFormula(string range, string value);
        //string RangeHAlign(string range, string value);
        //string RangeVAlign(string range, string value);
        //string RangeStyle(string range, string style, string value);
        //string RangeMerge(string range, string value);
        //string RangeImage(string patchImg, int colImg, int rowImg, int posTop, int posLeft, int imgWidth, int imgHeight);

    }
}
