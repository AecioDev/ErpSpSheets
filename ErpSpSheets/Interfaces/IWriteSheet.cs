using System;
using System.Runtime.InteropServices;

namespace ErpSpSheets
{
    [Guid("597E3623-207E-47FF-99E8-81F1F1560101"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), ComVisible(true)]
    public interface IWriteSheet
    {
        string OpenPlanilha(string FileName, string WorkSheet);
        string SavePlanilha();

        void CellValue(int ln, int col, string value);
        void CellFormula(int ln, int col, string value);
        void CellHAlign(int ln, int col, string value);
        void CellVAlign(int ln, int col, string value);
        void CellStyle(int ln, int col, string style, string value);

        void RangeValue(string range, string value);
        void RangeFormula(string range, string value);
        void RangeHAlign(string range, string value);
        void RangeVAlign(string range, string value);
        void RangeStyle(string range, string style, string value);
        void RangeMerge(string range, string value);
        void RangeImage(string patchImg, int colImg, int rowImg, int posTop, int posLeft, int imgWidth, int imgHeight);

    }
}
