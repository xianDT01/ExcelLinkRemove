using System;

namespace ExcelLinkRemover
{
    class Program
    {
        static void Main(string[] args)
        {
            EliminadorDeReferencias eliminador = new EliminadorDeReferencias();
            eliminador.ProcesarArchivosExcel();
        }
    }
}
