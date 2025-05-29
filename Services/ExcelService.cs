using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Drawing;
using System.Text.RegularExpressions;
using WebApplication1.Interfaces;
using WebApplication1.Models;


namespace WebApplication1.Services
{
    public class ExcelService : IExcelService
    {
        public Task<byte[]> ReturnExcelFile()
        {
            // Crea un nuevo libro de Excel en memoria
            using var workbook = new XLWorkbook();

            // Agrega una hoja llamada "DailyLogs"
            var worksheet = workbook.Worksheets.Add("DailyLogs");

          

            #region Text Principal
            // Combina las celdas desde la columna E hasta la I en la fila 5 (E5:I5)
            var mergedRange = worksheet.Range("E5:I5").Merge();

            // Establece el texto dentro de la celda combinada
            mergedRange.Value = "SERGIO MERCADO GONZALEZ";

            // Aplica un borde negro a todas las celdas del rango combinado
            mergedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            mergedRange.Style.Border.OutsideBorderColor = XLColor.Black;

            // Opcional: puedes centrar el texto dentro del rango
            mergedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            mergedRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            #endregion

            #region CONTROL DE TRAMITES PRESENTADOS AL INS  
            // Combina las celdas desde la columna E hasta la I en la fila 6 (E6:I6)
            var mergedRange2 = worksheet.Range("E6:I6").Merge();

            // Establece el texto dentro de la celda combinada
            mergedRange2.Value = "CONTROL DE TRAMITES PRESENTADOS AL INS";

            // Aplica un borde negro a todas las celdas del rango combinado
            mergedRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            mergedRange2.Style.Border.OutsideBorderColor = XLColor.Black;

            // Opcional: puedes centrar el texto dentro del rango
            mergedRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            mergedRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
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

            //
            #region
            var bitacoraCell = worksheet.Cell("L6");
            bitacoraCell.Value = "BITACORA";

            // Bordes negros alrededor de la celda
            bitacoraCell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            bitacoraCell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            bitacoraCell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            bitacoraCell.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            bitacoraCell.Style.Border.TopBorderColor = XLColor.Black;
            bitacoraCell.Style.Border.BottomBorderColor = XLColor.Black;
            bitacoraCell.Style.Border.LeftBorderColor = XLColor.Black;
            bitacoraCell.Style.Border.RightBorderColor = XLColor.Black;

            var cellK6 = worksheet.Cell("K6");
            cellK6.Value = ""; // Para que la celda exista
            cellK6.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            cellK6.Style.Border.RightBorderColor = XLColor.Black;

            // Centrar el texto
            bitacoraCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            bitacoraCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //negrilla del texto
            bitacoraCell.Style.Font.Bold = true;
            #endregion

            #region
            var cellM6 = worksheet.Cell("M6");
            cellM6.Value = "105-084";

            // Aplicar estilo
            cellM6.Style.Font.Bold = true;
            cellM6.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cellM6.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // Bordes negros
            cellM6.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            cellM6.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cellM6.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cellM6.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            cellM6.Style.Border.TopBorderColor = XLColor.Black;
            cellM6.Style.Border.BottomBorderColor = XLColor.Black;
            cellM6.Style.Border.LeftBorderColor = XLColor.Black;
            cellM6.Style.Border.RightBorderColor = XLColor.Black;

            //borde derecho een negro

            var cellN6 = worksheet.Cell("N6");
            cellN6.Value = ""; // Para que la celda exista
            cellN6.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cellN6.Style.Border.LeftBorderColor = XLColor.Black;

            #endregion

            #region
            // Establecer "Fecha" en L7 con estilo y bordes
            var cellL7 = worksheet.Cell("L7");
            cellL7.Value = "Fecha";

            // Negrilla
            cellL7.Style.Font.Bold = true;

            // Centrado horizontal y vertical
            cellL7.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cellL7.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // Borde negro completo para L7
            cellL7.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            cellL7.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cellL7.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cellL7.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            cellL7.Style.Border.TopBorderColor = XLColor.Black;
            cellL7.Style.Border.BottomBorderColor = XLColor.Black;
            cellL7.Style.Border.LeftBorderColor = XLColor.Black;
            cellL7.Style.Border.RightBorderColor = XLColor.Black;

            // También pintar el borde derecho de K7 (para que se note que está pegada a L7)
            var cellK7 = worksheet.Cell("K7");
            cellK7.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            cellK7.Style.Border.RightBorderColor = XLColor.Black;
            #endregion

            #region Fecha en M7

            // Establecer "28/02/2025" en M7
            var cellM7 = worksheet.Cell("M7");
            cellM7.Value = "28/02/2025";

            // Opcional: puedes formatear la celda como fecha
            cellM7.Style.DateFormat.Format = "dd/MM/yyyy";

            // Negrilla
            cellM7.Style.Font.Bold = true;

            // Centrado
            cellM7.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cellM7.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // Borde negro completo
            cellM7.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            cellM7.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cellM7.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cellM7.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            cellM7.Style.Border.TopBorderColor = XLColor.Black;
            cellM7.Style.Border.BottomBorderColor = XLColor.Black;
            cellM7.Style.Border.LeftBorderColor = XLColor.Black;
            cellM7.Style.Border.RightBorderColor = XLColor.Black;

            // Pintar borde izquierdo de la celda N7
            var cellN7 = worksheet.Cell("N7");
            cellN7.Value = " "; // Espacio en blanco
            cellN7.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cellN7.Style.Border.LeftBorderColor = XLColor.Black;

            #endregion

         

            #region
            // Escribir el texto
            string texto = "CHEQUE: 'C'; DEPOSITO: 'D'; EFECTIVO: 'E'; SINPE MOVIL: 'S'; TRANSFERENCIA: 'T'; VOUCHER´S: 'V'";
            var rango = worksheet.Range("C9:M9").Merge();

            // Establecer el valor
            rango.Value = texto;

            // Aplicar estilo: negrita y bordes
            rango.Style.Font.Bold = true;

            #endregion

            #region Colorear encabezado doble fila B10:P11
            var rangoColoreado = worksheet.Range("B10:P11");
            rangoColoreado.Style.Fill.BackgroundColor = XLColor.FromHtml("#b0c4de");
            #endregion

            #region 
            //b10
            
            var celdaNumero1 = worksheet.Cell("B10");
            celdaNumero1.Value = "NÚMERO";
            celdaNumero1.Style.Font.Bold = true;
            celdaNumero1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumero1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumero1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumero1.Style.Border.OutsideBorderColor = XLColor.Black;

            //B11
            var celdaNumero2 = worksheet.Cell("B11");
            celdaNumero2.Value = "REC / COM";
            celdaNumero2.Style.Font.Bold = true;
            celdaNumero2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumero2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumero2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumero2.Style.Border.OutsideBorderColor = XLColor.Black;
            

            //Formato celda C10
            var celdaNumero = worksheet.Cell("C10");
            celdaNumero.Value = "NÚMERO";
            celdaNumero.Style.Font.Bold = true;
            celdaNumero.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumero.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumero.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumero.Style.Border.OutsideBorderColor = XLColor.Black;

            //c11

            var celdaNumero3 = worksheet.Cell("C11");
            celdaNumero3.Value = "POLIZA";
            celdaNumero3.Style.Font.Bold = true;
            celdaNumero3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumero3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumero3.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumero3.Style.Border.OutsideBorderColor = XLColor.Black;

            //D10
            
            var celdaNumeroD10 = worksheet.Cell("D10");
            celdaNumeroD10.Value = "TIPO";
            celdaNumeroD10.Style.Font.Bold = true;
            celdaNumeroD10.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroD10.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroD10.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroD10.Style.Border.OutsideBorderColor = XLColor.Black;

            // D11

            var celdaNumeroD11 = worksheet.Cell("D11");
            celdaNumeroD11.Value = "SEGURO";
            celdaNumeroD11.Style.Font.Bold = true;
            celdaNumeroD11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroD11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroD11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroD11.Style.Border.OutsideBorderColor = XLColor.Black;

            //E10-F10

            var rangoEF10 = worksheet.Range("E10:F10").Merge();
            rangoEF10.Value = "VENCIMIENTO";
            rangoEF10.Style.Font.Bold = true;
            rangoEF10.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoEF10.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoEF10.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoEF10.Style.Border.InsideBorderColor = XLColor.Black;

            //centrar el textoef10
            rangoEF10.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoEF10.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            #endregion

            #region ajuste tamano de columnas

            //E11

            var celdaNumeroE10 = worksheet.Cell("E11");
            celdaNumeroE10.Value = "DE";
            celdaNumeroE10.Style.Font.Bold = true;
            celdaNumeroE10.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroE10.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroE10.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroE10.Style.Border.OutsideBorderColor = XLColor.Black;

            //F12

            var celdaNumeroF11 = worksheet.Cell("F11");
            celdaNumeroF11.Value = "HASTA";
            celdaNumeroF11.Style.Font.Bold = true;
            celdaNumeroF11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroF11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroF11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroF11.Style.Border.OutsideBorderColor = XLColor.Black;

            //G10-G11

            var rangog10_11 = worksheet.Range("G10:G11").Merge();
            rangog10_11.Value = "PRIMA";
            rangog10_11.Style.Font.Bold = true;
            rangog10_11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangog10_11.Style.Border.OutsideBorderColor = XLColor.Black;
            rangog10_11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangog10_11.Style.Border.InsideBorderColor = XLColor.Black;

            rangog10_11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangog10_11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

           
            //H10-H11

            var rangoH10_11 = worksheet.Range("H10:H11").Merge();
            rangoH10_11.Value = " ";
            rangoH10_11.Style.Font.Bold = true;
            rangoH10_11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoH10_11.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoH10_11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoH10_11.Style.Border.InsideBorderColor = XLColor.Black;

            rangoH10_11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoH10_11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //I10 

            var celdaNumeroI10 = worksheet.Cell("I10");
            celdaNumeroI10.Value = "MEDIO DE";
            celdaNumeroI10.Style.Font.Bold = true;
            celdaNumeroI10.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroI10.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroI10.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroI10.Style.Border.OutsideBorderColor = XLColor.Black;

            //i11

            var celdaNumeroI11 = worksheet.Cell("I11");
            celdaNumeroI11.Value = "PAGO";
            celdaNumeroI11.Style.Font.Bold = true;
            celdaNumeroI11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroI11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroI11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroI11.Style.Border.OutsideBorderColor = XLColor.Black;

            //J10

            var celdaNumeroJ10 = worksheet.Cell("J10");
            celdaNumeroJ10.Value = "TIPO";
            celdaNumeroJ10.Style.Font.Bold = true;
            celdaNumeroJ10.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroJ10.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroJ10.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroJ10.Style.Border.OutsideBorderColor = XLColor.Black;

            //J11

            var celdaNumeroJ11 = worksheet.Cell("J11");
            celdaNumeroJ11.Value = "TRANSACCION";
            celdaNumeroJ11.Style.Font.Bold = true;
            celdaNumeroJ11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            celdaNumeroJ11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            celdaNumeroJ11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            celdaNumeroJ11.Style.Border.OutsideBorderColor = XLColor.Black;

            //K10-K11
           

            var rangoK10_11 = worksheet.Range("K10:K11").Merge();
            rangoK10_11.Value = "ASEGURADO";
            rangoK10_11.Style.Font.Bold = true;
            rangoK10_11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoK10_11.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoK10_11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoK10_11.Style.Border.InsideBorderColor = XLColor.Black;

            rangoK10_11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoK10_11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //L10-L11

            var rangoL10_11 = worksheet.Range("L10:L11").Merge();
            rangoL10_11.Value = "PLACA";
            rangoL10_11.Style.Font.Bold = true;
            rangoL10_11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoL10_11.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoL10_11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoL10_11.Style.Border.InsideBorderColor = XLColor.Black;

            rangoL10_11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoL10_11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //M10-M11

            var rangoM10_11 = worksheet.Range("M10:M11").Merge();
            rangoM10_11.Value = "CEDULA";
            rangoM10_11.Style.Font.Bold = true;
            rangoM10_11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoM10_11.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoM10_11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoM10_11.Style.Border.InsideBorderColor = XLColor.Black;

            rangoM10_11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoM10_11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;


            //O10-O11
            var rangoO10_11 = worksheet.Range("O10:O11").Merge();
            rangoO10_11.Value = "TELEFONOS";
            rangoO10_11.Style.Font.Bold = true;
            rangoO10_11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoO10_11.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoO10_11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoO10_11.Style.Border.InsideBorderColor = XLColor.Black;

            rangoO10_11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoO10_11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //P10-P11
            var rangoP10_11 = worksheet.Range("P10:P11").Merge();
            rangoP10_11.Value = "CARGO MANUAL";
            rangoP10_11.Style.Font.Bold = true;
            rangoP10_11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoP10_11.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoP10_11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoP10_11.Style.Border.InsideBorderColor = XLColor.Black;

            rangoP10_11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangoP10_11.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //
            int filaInicioDaily = 12;

            foreach (var item in GetDailyLogs()
             .Where(item => item.TransactionType != "SOBRANTE" &&
                            item.TransactionType != "AMPLIACION DE VIGENCIA RT"))
            {
                worksheet.Cell(filaInicioDaily, "B").Value = item.ReceiptNumber;
                worksheet.Cell(filaInicioDaily, "C").Value = item.Code.StartsWith("EMISION") ? " " : item.Code;
                worksheet.Cell(filaInicioDaily, "D").Value = item.InsuranceType;
                worksheet.Cell(filaInicioDaily, "E").Value = item.From.ToString("dd/MM/yyyy");
                worksheet.Cell(filaInicioDaily, "F").Value = item.To.ToString("dd/MM/yyyy");
                worksheet.Cell(filaInicioDaily, "G").Value = item.Amount;
                
                // Obtener la primera cifra del monto decimal como string
                var PaymentMethodStr = item.PaymentMethod.ToString();
                worksheet.Cell(filaInicioDaily, "H").Value =
                !string.IsNullOrEmpty(PaymentMethodStr) ? PaymentMethodStr[0].ToString() : "";

             
                // Obtener solo los números de item.Amount (ej. "V-0988" -> "0988")
                string amountCode = item.PaymentMethod; // Asegúrate de que item.Amount sea string
                string numericOnly = Regex.Replace(amountCode ?? "", @"[^\d]", "");

                worksheet.Cell(filaInicioDaily, "I").Value = numericOnly;

                worksheet.Cell(filaInicioDaily, "J").Value = item.TransactionType;
                //

                worksheet.Cell(filaInicioDaily, "K").Value = item.InsuredFullName;
                worksheet.Cell(filaInicioDaily, "L").Value = item.LicensePlate;
                worksheet.Cell(filaInicioDaily, "M").Value = item.Identification;
                worksheet.Cell(filaInicioDaily, "O").Value = string.IsNullOrEmpty(item.PhoneNumber) ? string.Empty : item.PhoneNumber;
                worksheet.Cell(filaInicioDaily, "N").Value = string.IsNullOrEmpty(item.ChargeDay) ? string.Empty : item.ChargeDay;

                filaInicioDaily++;
            }
                            

            var rangoPrueba = worksheet.Range($"P{filaInicioDaily}:P{filaInicioDaily + 1}").Merge();
            rangoPrueba.Value = "CARGO MANUAL";
            rangoPrueba.Style.Font.Bold = true;

            #endregion
            #region Auto‑ajuste de ancho de columnas
            // Ajusta el ancho de TODAS las columnas que tienen contenido
            worksheet.ColumnsUsed().AdjustToContents();
// Si quieres abarcar incluso aquellas columnas que no aparecen en ColumnsUsed():
// worksheet.Columns(1, worksheet.LastColumnUsed().ColumnNumber()).AdjustToContents();
#endregion
            // Guardar el libro en un stream de memoria
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);

            // Retornar el contenido del archivo como arreglo de bytes
            return Task.FromResult(stream.ToArray());


            //foreach (var item in GetDailyLogs())
            //{

            //}

            //foreach (var item in GetProcedures())
            //{

            //}

            
        }


        private List<DailyLog> GetDailyLogs()
        {
            return new List<DailyLog>()
            {
                        //TODO: Crear los que faltan
        new DailyLog("343712","0104AUT0258111-01","AUTOS",new DateTime(2025,02,19),new DateTime(2025,03,19),63140,"V-508","RENOVACION","BELLIDO SOLANO LUIS JOSE","402440256",DateTime.Now,"50661393498","CL289812"),
        new DailyLog("343890","0104AUT0262031-00","AUTOS",new DateTime(2025,02,28),new DateTime(2025,03,30),32793,"V-515","RENOVACION","CHAVARRIA MORA JORGE STEVEN","113480170",DateTime.Now,"50688681066","BHL565","30"),
        new DailyLog("343770","0104AUT0170751-20","AUTOS",new DateTime(2025,02,21),new DateTime(2025,03,21),19615,"V-516","RENOVACION","BOLAÑOS CHAVES ANGIE TATIANA","111300987",DateTime.Now,"50662520001","BGP915","30"),
        new DailyLog("227831","0104AUT0264122-00","AUTOS",new DateTime(2025,02,28),new DateTime(2025,05,29),191552,"V-517","RENOVACION","BALMACEDA ARAGON ALEJANDRO","109760188",DateTime.Now,"50688587203","PLAN FAM VIQUEZ CRUZ MARCO VINICIO BDG307"),
        new DailyLog("346141","0104AUT0211279-13","AUTOS",new DateTime(2025,02,14),new DateTime(2025,03,14),21185,"V-519","RENOVACION","VIQUEZ BARQUERO YAMILETH","401390032",DateTime.Now,"50683635256","401753"),
        new DailyLog("292133","0104AUT0254422-02","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),35432,"V-520","RENOVACION","RAMOS GONZALEZ MELISSA","114020771",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO MRG021"),
        new DailyLog("292143","0104AUT0257398-01","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),22259,"V-520","RENOVACION","ROJAS ALVAREZ RODOLFO","105250453",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO BVP240"),
        new DailyLog("292158","0104AUT0242919-06","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),67099,"V-520","RENOVACION","ROJAS HIDALGO PABLO","111600202",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO GGG825"),
        new DailyLog("256403","0104AUT0255152-02","AUTOS",new DateTime(2025,02,15),new DateTime(2025,05,15),24400,"S-520","RENOVACION","ROJAS HIDALGO PABLO","111600202",DateTime.Now,"50683303503","74000"),
        new DailyLog("488596","0104AUT0265129-00","AUTOS",new DateTime(2025,02,15),new DateTime(2025,03,15),25320,"V-520","RENOVACION","ROJAS HIDALGO PABLO","111600202",DateTime.Now,"50683303503","PLAN FAM ROJAS HIDALGO PABLO AAG575"),
        new DailyLog("346039","0104AUT0251890-03","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),8665,"V-521","RENOVACION","SMITH SAENZ JADELYN SHADAY","402620232",DateTime.Now,"50685861010","875298"),
        new DailyLog("227751","0104AUT0257544-01","AUTOS",new DateTime(2025,02,14),new DateTime(2025,05,14),51742,"V-522","RENOVACION","ARIAS TORRENTES ANDREY DAVID","603470399",DateTime.Now,"50683028494","PLAN FAM ARIAS TORRENTES ANDREY DAVID BSM604"),
        new DailyLog("320210","0104AUT0266498-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),31979,"V-523","RENOVACION","GONZALEZ PEREZ JORGE ENRIQUE","602000204",DateTime.Now,"50688527020","CL313750"),
        new DailyLog("376302","0104AUT0267210-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),41526,"V-523","RENOVACION","GONZALEZ PEREZ JORGE ENRIQUE","602000204",DateTime.Now,"50688527020","CL349838"),
        new DailyLog("345359","0104AUT0264241-00","AUTOS",new DateTime(2025,02,28),new DateTime(2025,03,31),32403,"S-968620","RENOVACION","MADRIGAL LIZANO MARIA CELINA","108440638",DateTime.Now,"50687046532","PLAN FAM MADRIGAL LIZANO MARIA CELINA LST375"),
        new DailyLog("346101","0104AUT0249724-04","AUTOS",new DateTime(2025,02,18),new DateTime(2025,03,18),22220,"T-10433778","RENOVACION","TREJOS MENDEZ CAROLINA","116450108",DateTime.Now,"50684271114","802538"),
        new DailyLog("292040","0104AUT0256985-01","AUTOS",new DateTime(2025,02,14),new DateTime(2025,03,14),66023,"T-11271029","RENOVACION","MONGE ALFARO ALLAN","401580701",DateTime.Now,"50689807948","PLAN FAM MONGE ALFARO ALLAN RRS421"),
        new DailyLog("395161","0104IMR0007746-00","INCENDIO",new DateTime(2025,08,03),new DateTime(2025,08,04),13185,"E-EFECTIVO","RENOVACION","GUZMAN SOLANO DIRSEO GERARDO","111580425",DateTime.Now,"50685830037"),
        new DailyLog("400997","0104AUT0221549-12","AUTOS",new DateTime(2025,08,03),new DateTime(2025,08,04),14149,"E-EFECTIVO","RENOVACION","MERCADO GONZALEZ BYRON ANTONIO","155808624707",DateTime.Now,"50687128320","BQQ382"),
        new DailyLog("396475","0104IMR0006490-01","HOGAR COMPRENSIVO ",new DateTime(2025,09,03),new DateTime(2026,09,03),67949,"V-523","RENOVACION","MORA TORRES KATHERINE","114080036",DateTime.Now,"50661242278"),
        new DailyLog("488233","0104AUT0267097-00","AUTOS",new DateTime(2025,03,03),new DateTime(2025,03,06),140286,"V-518","RENOVACION","3 101 882035 S A","3101882035",DateTime.Now,"50670339600","DSY069"),
        new DailyLog("345134","0104AUT0261820-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),27090,"V-524","RENOVACION","HERRERA JIMENEZ JOHNNATAN ALEJANDRO","401930721",DateTime.Now,"50688095858","PMK090"),
        new DailyLog("345165","0104AUT0261801-00","AUTOS",new DateTime(2025,02,16),new DateTime(2025,03,16),20534,"V-524","RENOVACION","HERRERA JIMENEZ JOHNNATAN ALEJANDRO","401930721",DateTime.Now,"50688095858","BYD260"),
        new DailyLog("5612974"," ","AUTOS",new DateTime(2025,02,28),new DateTime(2025,03,28),21260,"S-734311","EMISION","VENEGAS CAMPOS VERONICA","110790786",DateTime.Now,"50689795454","C151370"),
        new DailyLog("293912","D006177","VIDA",new DateTime(2025,02,16),new DateTime(2025,03,16),25430,"V-508","RENOVACION","BELLIDO VARGAS JOSE LUIS","502820347",DateTime.Now),
        new DailyLog("343741","VG4220020040731","VIDA GLOBAL ",new DateTime(2025,02,16),new DateTime(2025,03,16),13727,"V-508","EMISION","BELLIDO VARGAS JOSE LUIS","502820347",DateTime.Now),
        new DailyLog("487177","0104ACI0275987-00","ESTUDIANTIL",new DateTime(2025,02,28),new DateTime(2026,02,28),20430,"E-EFECTIVO","EMISION","SOLANO GARRO YANCY MAGALLY","110670140",DateTime.Now,"50685751144")
            };
        }

        private List<ProcedureWithoutMoney> GetProcedures()
        {
            return new List<ProcedureWithoutMoney>()
            {
               new ProcedureWithoutMoney("0204VIA0087142-00","RODRIGUEZ RODRIGUEZ JORGE EDUARDO","CARTA SOLICITUD RESERVA DE DINERO")
            };
        }
    }
}