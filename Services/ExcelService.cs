using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Net;
using System.Text.RegularExpressions;
using WebApplication1.Interfaces;
using WebApplication1.Models;


namespace WebApplication1.Services
{


    public static class DailyLogExtensions
    {
        public static decimal GetSum(this List<DailyLog> dailylogs, string initial)
        {
            return dailylogs
                .Where(x => x.PaymentMethod.StartsWith(initial))
                .Sum(x => x.Amount);
        }

        public static void SetHeader(this IXLWorksheet ws, string cell, string value)
        {
            var c = ws.Cell(cell);
            c.Value = value;
            c.ApplyHeaderStyle();
        }
        public static void ApplyHeaderStyle(this IXLCell cell)
        {
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.OutsideBorderColor = XLColor.Black;
        }

        public static void ApplyHeaderStyle(this IXLRange range)
        {
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.OutsideBorderColor = XLColor.Black;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorderColor = XLColor.Black;
        }

        public static void SetMergedHeader(this IXLWorksheet ws, string range, string value)
        {
            var merged = ws.Range(range).Merge();
            merged.Value = value;
            merged.ApplyHeaderStyle();
        }
    }

    public class ExcelService : IExcelService
    {
        public Task<byte[]> ReturnExcelFile(IEnumerable<DailyLog> dailyLogs, IEnumerable<ProcedureWithoutMoney> procedures, string dailyLogNumber)
        {
            // Crea un nuevo libro de Excel en memoria
            using var workbook = new XLWorkbook();
            // Agrega una hoja llamada "DailyLogs"
            var worksheet = workbook.Worksheets.Add("DailyLogs");
            #region Text Principal
            // SERGIO M G
            worksheet.SetMergedHeader("E5:I5", "SERGIO MERCADO GONZALEZ");
            #endregion

            #region CONTROL DE TRAMITES PRESENTADOS AL INS  
            worksheet.SetMergedHeader("E6:I6", "CONTROL DE TRAMITES PRESENTADOS AL INS");
            #endregion

            #region
            double heightB6 = worksheet.Row(6).Height;
            double heightB7 = worksheet.Row(7).Height;
            double heightB8 = worksheet.Row(8).Height;

            // Convertir a píxeles (1 punto ≈ 0.75 píxeles)
            int totalHeightPx = (int)((heightB6 + heightB7 + heightB8) * 0.75);

            // Ancho en píxeles (columna B)
            double widthB = worksheet.Column("B").Width;
            int widthPx = (int)(widthB * 7); // ClosedXML usa ~7px por unidad de ancho

            var imagePath = "C:\\Users\\Danilo\\source\\repos\\WebApplication1\\WebApplication1\\NewFolder1\\faviconc.png";

            if (File.Exists(imagePath))
            {
                using var imageStream = File.OpenRead(imagePath);

                var picture = worksheet.AddPicture(imageStream)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .WithSize(widthPx, totalHeightPx) // Imagen exactamente del alto de B6:B8
                    .MoveTo(worksheet.Cell("B6"), 0, 0); // Desde la esquina superior de B6
            }
            #endregion
            #region
            //L6

            worksheet.SetHeader("L6", "BITACORA");

            //M6
            worksheet.SetHeader("M6", $"105-{dailyLogNumber}");

            // Establecer "Fecha" en L7 con estilo y bordes
            worksheet.SetHeader("L7", "Fecha");

            // Establecer "28/02/2025" en M7
            worksheet.SetHeader("M7", "28/02/2025");
            #endregion

            #region
            // Esto no lleva bordes por lo cual no cabe en la funcion
            string texto = "CHEQUE: 'C'; DEPOSITO: 'D'; EFECTIVO: 'E'; SINPE MOVIL: 'S'; TRANSFERENCIA: 'T'; VOUCHER´S: 'V'";
            var rango = worksheet.Range("C9:M9").Merge();
            rango.Value = texto;
            rango.Style.Font.Bold = true;
            #endregion

            #region Colorear encabezado doble fila B10:P11
            var rangoColoreado = worksheet.Range("B10:O11");
            rangoColoreado.Style.Fill.BackgroundColor = XLColor.FromHtml("#b0c4de");
            #endregion

            #region 
            //b10
            worksheet.SetHeader("B10", "NÚMERO");
            //B11
            worksheet.SetHeader("B11", "REC / COM");
            // C10
            worksheet.SetHeader("C10", "NÚMERO");
            //c11
            worksheet.SetHeader("C11", "POLIZA");
            //D10
            worksheet.SetHeader("D10", "TIPO");
            // D11
            worksheet.SetHeader("D11", "SEGURO");
            //E10-F10
            worksheet.SetMergedHeader("E10:F10", "VENCIMIENTO");
            #endregion
            #region ajuste tamano de columnas

            //E11
            worksheet.SetHeader("E11", "DE");

            //F11
            worksheet.SetHeader("F11", "HASTA");

            //G10-G11
            worksheet.SetMergedHeader("G10:G11", "PRIMA");

            //H10-H11
            worksheet.SetMergedHeader("H10:H11", " ");

            //I10 
            worksheet.SetHeader("I10", "MEDIO DE");

            //i11
            worksheet.SetHeader("I11", "PAGO");

            //J10
            worksheet.SetHeader("J10", "TIPO");

            //J11
            worksheet.SetHeader("J11", "TRANSACCION");

            //K10-K11
            worksheet.SetMergedHeader("K10:K11", "ASEGURADO");

            //L10-L11
            worksheet.SetMergedHeader("L10:L11", "PLACA");

            //M10-M11
            worksheet.SetMergedHeader("M10:M11", "CEDULA");

            //n10-O11
            worksheet.SetMergedHeader("N10:N11", "TELEFONOS");

            //O10-P11
            worksheet.SetMergedHeader("O10:O11", "CARGO MANUAL");


            // EN ESTA FILA EMPIEZA LODINAMICO PORQUE EL RESTO ES QUEMADO
            int filaInicioDaily = 12;
            decimal totalColones = 0; // Aquí se guardará la suma de todos los montos en colones

            foreach (var item in dailyLogs
                .Where(item => item.TransactionType != "SOBRANTE" &&
                               item.TransactionType != "AMPLIACION DE VIGENCIA RT"))
            {
                worksheet.Cell(filaInicioDaily, "B").Value = item.ReceiptNumber;
                worksheet.Cell(filaInicioDaily, "C").Value = item.Code.StartsWith("EMISION") ? " " : item.Code;
                worksheet.Cell(filaInicioDaily, "D").Value = item.InsuranceType;
                worksheet.Cell(filaInicioDaily, "E").Value = item.From.ToString("dd/MM/yyyy");
                worksheet.Cell(filaInicioDaily, "F").Value = item.To.ToString("dd/MM/yyyy");
                var PaymentMethodStr = item.PaymentMethod.ToString();
                worksheet.Cell(filaInicioDaily, "H").Value =
                !string.IsNullOrEmpty(PaymentMethodStr) ? PaymentMethodStr[0].ToString() : "";

                string amountCode = item.PaymentMethod;
                string numericOnly = Regex.Replace(amountCode ?? "", @"[^\d]", "");
                worksheet.Cell(filaInicioDaily, "I").Value = numericOnly;
                worksheet.Cell(filaInicioDaily, "J").Value = item.TransactionType;
                worksheet.Cell(filaInicioDaily, "K").Value = item.InsuredFullName;
                worksheet.Cell(filaInicioDaily, "L").Value = item.LicensePlate;
                worksheet.Cell(filaInicioDaily, "M").Value = item.Identification;
                worksheet.Cell(filaInicioDaily, "N").Value = string.IsNullOrEmpty(item.PhoneNumber) ? string.Empty : item.PhoneNumber;
                worksheet.Cell(filaInicioDaily, "O").Value = item.ChargeDay;

                // Aplicar borde negro a toda la fila de la B a la O
                var range = worksheet.Range($"B{filaInicioDaily}:O{filaInicioDaily}");
                range.ApplyHeaderStyle();

                // Columna G - Monto en colones
                var celdaColones = worksheet.Cell(filaInicioDaily, "G");
                celdaColones.Value = item.Amount;
                celdaColones.Style.NumberFormat.Format = "₡ #,##0.00";

                // Acumular el monto
                totalColones += item.Amount;
                filaInicioDaily++;
            }

            // Escribir el total en la celda justo debajo de la última fila en la columna G
            var celdaTotalColones = worksheet.Cell(filaInicioDaily, "G");
            celdaTotalColones.Value = totalColones;
            celdaTotalColones.Style.Font.Bold = true;
            celdaTotalColones.Style.NumberFormat.Format = "₡ #,##0.00";
            celdaTotalColones.Style.Fill.BackgroundColor = XLColor.Orange;

            // Aplicar bordes negros a toda la fila (de B a P)
            var rangoFilaExtra = worksheet.Range($"B{filaInicioDaily}:O{filaInicioDaily}");
            rangoFilaExtra.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoFilaExtra.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoFilaExtra.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoFilaExtra.Style.Border.InsideBorderColor = XLColor.Black;

            // Lista de valores que irán en la columna B
            var conceptos = new List<(string, XLColor)>
            {   ("CHEQUES:", XLColor.FromArgb(191, 191, 191)),
                ("DEPOSITOS:", XLColor.FromArgb(255, 255, 255)),
                ("EFECTIVO:", XLColor.FromArgb(191, 191, 191)),
                ("SINPE MOVIL:", XLColor.FromArgb(255, 255, 255)),
                ("TRANSF. ELECT.", XLColor.FromArgb(191, 191, 191)),
                ("VOUCHER´S:", XLColor.FromArgb(255, 255, 255))
            };

            // Empezamos desde una fila debajo
            int filaBase = filaInicioDaily;
            // Recorrer cada concepto y dibujar fila con estilo
            foreach (var (text, color) in conceptos)
            { //
                var celdaTexto = worksheet.Cell(filaBase, "B");
                celdaTexto.Value = text;
                celdaTexto.Style.Font.Bold = true;

                var celdaValor = worksheet.Cell(filaBase, "C");
                celdaTexto.Style.Fill.BackgroundColor = color;
                celdaValor.Style.Fill.BackgroundColor = color;

                celdaValor.Value = dailyLogs.ToList().GetSum(text[..1]);
                celdaValor.Style.NumberFormat.Format = "₡ #,##0.00";


                var rangoFila = worksheet.Range($"B{filaBase}:C{filaBase}");
                rangoFila.ApplyHeaderStyle();

                filaBase++; // Avanzar a la siguiente fila
               
            }

        


            // Dos filas con tres celdas (B, C, D)
            var valores = new List<(string B, string C, string D)>
                     {
                         ("POLIZA", "NOMBRE", "TRAMITE")
                    };

            foreach (var (B, C, D) in valores)
            {
                worksheet.Cell(filaBase, "B").Value = B;
                worksheet.Cell(filaBase, "C").Value = C;
                worksheet.Cell(filaBase, "D").Value = D;

                var range = worksheet.Range($"B{filaBase}:D{filaBase}");
                range.ApplyHeaderStyle();
                range.Style.Fill.BackgroundColor = XLColor.FromHtml("#b0c4de"); // Fondo azul claro
                range.Style.Font.Bold = true; // Negrita

                filaBase++;
            }

            foreach (var item in procedures)
            {
                worksheet.Cell(filaBase, "B").Value = item.InsuranceCode;
                worksheet.Cell(filaBase, "C").Value = item.Insured;
                worksheet.Cell(filaBase, "D").Value = item.Procedure;

                var range2 = worksheet.Range($"B{filaBase}:D{filaBase}");
                range2.ApplyHeaderStyle();

                filaBase++;
            }

            #endregion
            #region Auto‑ajuste de ancho de columnas
            // Ajusta el ancho de TODAS las columnas que tienen contenido
            worksheet.ColumnsUsed().AdjustToContents();
            worksheet.Column("G").Width = 15;
            worksheet.Column("O").Width = 15;
            // Si quieres abarcar incluso aquellas columnas que no aparecen en ColumnsUsed():
            // worksheet.Columns(1, worksheet.LastColumnUsed().ColumnNumber()).AdjustToContents();
            #endregion
            // Guardar el libro en un stream de memoria
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);

            // Retornar el contenido del archivo como arreglo de bytes
            return Task.FromResult(stream.ToArray());

        }

    }
}