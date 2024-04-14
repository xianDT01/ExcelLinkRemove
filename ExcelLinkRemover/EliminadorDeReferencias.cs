using System;
using System.IO;

namespace ExcelLinkRemover
{
    class EliminadorDeReferencias
    {
        public void ProcesarArchivosExcel()
        {
            string directorioPath = SolicitarDirectorio();
            try
            {
                string[] archivosExcel = Directory.GetFiles(directorioPath, "*.xlsm");
                foreach (string archivo in archivosExcel)
                {
                    EliminarReferenciasExternas(archivo);
                    Console.WriteLine($"Vínculos de libro eliminados en {archivo}");
                }
                Console.WriteLine("Proceso de la eliminación de Vínculos finalizado");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private string SolicitarDirectorio()
        {
            Console.WriteLine("Introduce la ruta del directorio:");
            string directorioPath = Console.ReadLine();

            if (!Directory.Exists(directorioPath))
            {
                Console.WriteLine("El directorio especificado no existe");
            }

            return directorioPath;
        }

        private void EliminarReferenciasExternas(string filePath)
        {
            using (var spreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, true))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;

                foreach (var externalWorkbookPart in workbookPart.ExternalWorkbookParts)
                {
                    workbookPart.DeletePart(externalWorkbookPart);
                }

                workbookPart.Workbook.Save();
            }
        }
    }
}

