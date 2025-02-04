using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace AplicacionExcelsTomi
{
    public class Program
    {
        static void Main(string[] args)
        {
            // Aceptar la licencia de EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string carpetaMain = @"D:\PLANILLAS_TOMI\"; // Harcodeado

            List<string[]> datosExtraidos = new List<string[]>();
            datosExtraidos.Add(new string[] { "NOMBRE Y APELLIDO", "WPP/CELULAR", "COLEGIO Y CURSO" }); // Encabezado

            List<string> carpetas = Directory.EnumerateDirectories(carpetaMain).ToList();
            
            string carpetaExcel;

            for (int i = 0; i < carpetas.Count(); i++)
            {
                carpetaExcel = carpetas[i];
                char splitter = '\\'; 
                string archivoSalida = carpetas[i] + ".csv"; //D:\PLANILLAS_TOMI\CABALLITO-PALERMO\
                string archivoXlsx = carpetas[i] + ".xlsx";

                foreach (string archivo in Directory.GetFiles(carpetaExcel, "*.xlsx"))
                {
                    try
                    {
                        // Evitar archivos temporales (~$)
                        if (Path.GetFileName(archivo).StartsWith("~$"))
                        {
                            //Console.WriteLine($"Ignorando archivo temporal: {archivo}");
                            continue;
                        }

                        using (var package = new ExcelPackage(new FileInfo(archivo)))
                        {
                            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                            if (worksheet == null) continue;

                            int colNombre = 0, colTelefono = 0;
                            int totalColumnas = worksheet.Dimension?.Columns ?? 0;
                            int totalFilas = worksheet.Dimension?.Rows ?? 0;

                            if (totalColumnas == 0 || totalFilas == 0) continue; // Evitar archivos vacíos

                            // Buscar las columnas de Nombre y Teléfono
                            for (int col = 1; col <= totalColumnas; col++)
                            {
                                string header = worksheet.Cells[2, col].Text.Trim().ToLower();
                                if (header == "nombre y apellido") colNombre = 2;
                                if (header == "wpp/celular*" || header == "telefono") colTelefono = 4;
                            }

                            if (colNombre == 0 || colTelefono == 0) continue; // Saltar si no encuentra las columnas

                            string colegio = worksheet.Cells[1, 1].Text.Trim().Replace("COLEGIO Y CURSO","").Replace(":", "");

                            // Leer los datos
                            for (int fila = 4; fila <= totalFilas; fila++)
                            {
                                string nombre = worksheet.Cells[fila, colNombre].Text.Trim().Replace(",","");
                                string telefono = worksheet.Cells[fila, colTelefono].Text.Trim();
                                if (!string.IsNullOrEmpty(nombre) || !string.IsNullOrEmpty(telefono))
                                {
                                    datosExtraidos.Add(new string[] { nombre, telefono, colegio});
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al procesar {archivo}: {ex.Message}");
                    }
                }    

            // Guardar en CSV
            System.IO.File.WriteAllLines(archivoSalida, datosExtraidos.Select(d => string.Join(",", d)));
            //Console.WriteLine($"Datos extraídos correctamente en {archivoSalida}");
            // Convertir a xlsx
            ConvertCsvToXlsx(archivoSalida, archivoXlsx);
            System.IO.File.Delete(archivoSalida);
            Console.WriteLine($"XLSX creado en {archivoXlsx}");
            datosExtraidos = new List<string[]>();
            datosExtraidos.Add(new string[] { "NOMBRE Y APELLIDO", "WPP/CELULAR", "COLEGIO Y CURSO" }); // Encabezado
            }
        }

        public static void ConvertCsvToXlsx(string csvPath, string xlsxPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Requerido para usar EPPlus gratis

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Hoja1");

                string[] lines = System.IO.File.ReadAllLines(csvPath);
                for (int row = 0; row < lines.Length; row++)
                {
                    string[] columns = lines[row].Split(','); // Ajusta el delimitador si es necesario
                    for (int col = 0; col < columns.Length; col++)
                    {
                        worksheet.Cells[row + 1, col + 1].Value = columns[col]; // Excel usa índices basados en 1
                    }
                }

                package.SaveAs(new FileInfo(xlsxPath));
            }
        }

    }
}
