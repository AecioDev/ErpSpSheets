using System;
using System.Runtime.InteropServices;

namespace ErpSpSheets
{
    [Guid("9CA006F9-50BD-4B7F-90F1-575127DDC2F0"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), ComVisible(true)]
    public interface IReadSheet
    {
        string OpenPlanilha(string FileName, string WorkSheet);
        string ReadText(int ln, int col);
        int ReadNumber(int ln, int col);
        decimal ReadReal(int ln, int col);
        DateTime ReadDate(int ln, int col);
        string ClosePlanilha();

    }
}
