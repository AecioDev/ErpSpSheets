using ErpSpSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TesteErpSheets
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadSheet ex = new ReadSheet();

            var resp = ex.OpenPlanilha(@"C:\Temp\PlanOrc.xlsx", "");
            if (string.IsNullOrEmpty(resp)) //Sem Erro
            {
                var data = ex.ReadDate(4, 4);
                var tipo1 = ex.ReadType(4, 4);

                var texto = ex.ReadText(6, 3);
                var tipo2 = ex.ReadType(6, 3);

                var numero = ex.ReadNumber(6, 2);
                var tipo3 = ex.ReadType(6, 2);

                var valor = ex.ReadReal(6, 4);
                var tipo4 = ex.ReadType(6, 4);

                Console.WriteLine("Data: " + data + ", Nome: " + texto + ", Código: " + numero + ", Valor: " + valor.ToString("C2") + "\n\n");

                Console.WriteLine("Tipo 1"+ tipo1.ToString()+ ", Tipo 2" + tipo2.ToString() + ", Tipo 3" + tipo3.ToString() + ", Tipo 4" + tipo4.ToString());
            }
            else
            {
                Console.WriteLine(resp);
            }

            ex.ClosePlanilha();

            Console.ReadLine();
        }
    }
}
