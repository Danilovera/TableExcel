using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Drawing;
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
    }

    public class ExcelService : IExcelService
    {
        void SetHeader(IXLWorksheet ws, string cell, string value)
        {
            var c = ws.Cell(cell);
            c.Value = value;
            ApplyHeaderStyle(c);
        }

        private void ApplyHeaderStyle(IXLCell cell)
        {
           // cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.OutsideBorderColor = XLColor.Black;

        }

        private void ApplyHeaderStyle(IXLRange range)
        {
           // range.Style.Font.Bold = true;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.OutsideBorderColor = XLColor.Black;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorderColor = XLColor.Black;
        
        }

        void SetMergedHeader(IXLWorksheet ws, string range, string value)
        {
            var merged = ws.Range(range).Merge();
            merged.Value = value;
            ApplyHeaderStyle(merged);
        }

        public Task<byte[]> ReturnExcelFile()
        {
            // Crea un nuevo libro de Excel en memoria
            using var workbook = new XLWorkbook();
            // Agrega una hoja llamada "DailyLogs"
            var worksheet = workbook.Worksheets.Add("DailyLogs");
            #region Text Principal
            // SERGIO M G
            SetMergedHeader(worksheet, "E5:I5", "SERGIO MERCADO GONZALEZ");
            #endregion

            #region CONTROL DE TRAMITES PRESENTADOS AL INS  
            SetMergedHeader(worksheet, "E6:I6", "CONTROL DE TRAMITES PRESENTADOS AL INS");
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
           
            SetHeader(worksheet, "L6", "BITACORA");

            //M6
            SetHeader(worksheet, "M6","105-084");
         
            // Establecer "Fecha" en L7 con estilo y bordes
            SetHeader(worksheet, "L7", "Fecha");

            // Establecer "28/02/2025" en M7
            SetHeader(worksheet, "M7", "28/02/2025");
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
            SetHeader(worksheet, "B10", "NÚMERO");
            //B11
            SetHeader(worksheet, "B11", "REC / COM");
            // C10
            SetHeader(worksheet, "C10", "NÚMERO");
            //c11
            SetHeader(worksheet, "C11", "POLIZA");
            //D10
            SetHeader(worksheet, "D10", "TIPO");
            // D11
            SetHeader(worksheet, "D11", "SEGURO");
            //E10-F10
            SetMergedHeader(worksheet, "E10:F10", "VENCIMIENTO");
            #endregion
            #region ajuste tamano de columnas

            //E11
            SetHeader(worksheet, "E11", "DE");

            //F12
            SetHeader(worksheet, "f11", "HASTA");

            //G10-G11
            SetMergedHeader(worksheet, "G10:G11", "PRIMA");

            //H10-H11
            SetMergedHeader(worksheet, "H10:H11", " ");

            //I10 
            SetHeader(worksheet, "I10","MEDIO DE");

            //i11
            SetHeader(worksheet, "I11", "PAGO");

            //J10
            SetHeader(worksheet, "J10", "TIPO");

            //J11
            SetHeader(worksheet, "J11", "TRANSACCION");

            //K10-K11
            SetMergedHeader(worksheet, "K10:K11", "ASEGURADO");

            //L10-L11
            SetMergedHeader(worksheet, "L10:L11", "PLACA");

            //M10-M11
            SetMergedHeader(worksheet, "M10:M11", "CEDULA");

            //n10-O11
            SetMergedHeader(worksheet, "N10:N11", "TELEFONOS");

            //O10-P11
            SetMergedHeader(worksheet, "O10:O11", "CARGO MANUAL");


            //
            int filaInicioDaily = 12;
            decimal totalColones = 0; // Aquí se guardará la suma de todos los montos en colones

            foreach (var item in GetDailyLogs()
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
                ApplyHeaderStyle(range);

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

            // Aplicar bordes negros a toda la fila (de B a P)
            var rangoFilaExtra = worksheet.Range($"B{filaInicioDaily}:O{filaInicioDaily}");
            rangoFilaExtra.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangoFilaExtra.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangoFilaExtra.Style.Border.OutsideBorderColor = XLColor.Black;
            rangoFilaExtra.Style.Border.InsideBorderColor = XLColor.Black;

            // Lista de valores que irán en la columna B
            var conceptos = new List<string>
            {   "CHEQUES:",
                "DEPOSITOS:",
                "EFECTIVO:",
                "SINPE MOVIL:",
                "TRANSF. ELECT.",
                "VOUCHER´S:"
            };

            // Empezamos desde una fila debajo"
            int filaBase = filaInicioDaily ;
            var dailyLogs = GetDailyLogs(); 

            // Recorrer cada concepto y dibujar fila con estilo
            foreach (var text in conceptos)
            {
                var celdaTexto = worksheet.Cell(filaBase, "B");
                celdaTexto.Value = text;
                celdaTexto.Style.Font.Bold = true;

                var celdaValor = worksheet.Cell(filaBase, "C");

                if (text == "DEPOSITOS:" || text == "TRANSF. ELECT.")
                {
                    celdaTexto.Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);
                    celdaValor.Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);
                }

                celdaValor.Value = dailyLogs.GetSum(text[..1]);

                var rangoFila = worksheet.Range($"B{filaBase}:C{filaBase}");
                rangoFila.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rangoFila.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rangoFila.Style.Border.OutsideBorderColor = XLColor.Black;
                rangoFila.Style.Border.InsideBorderColor = XLColor.Black;

                filaBase++; // Avanzar a la siguiente fila

                // Insertar después de "VOUCHER´S:"
                if (text == "VOUCHER´S:")
                {
                    // Fila con una celda en B
                    var celdaExtra1 = worksheet.Cell(filaBase, "B");
                    celdaExtra1.Value = "TRAMITES SIN DINERO";
                    celdaExtra1.Style.Font.Bold = true;
                    celdaExtra1.Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);
                    celdaExtra1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    filaBase++;

                    if (text == "CHEQUES:")
                    {
                        decimal sumaCheques = dailyLogs.GetSum("C");
                        celdaValor.Value = sumaCheques;
                        celdaValor.Style.NumberFormat.Format = "₡ #,##0.00";
                    }

                    // Dos filas con tres celdas (B, C, D)
                    var valores = new List<(string B, string C, string D)>
                     {
                         ("POLIZA", "NOMBRE", "TRAMITE"),
                         ("0204VIA0087142-00", "RODRIGUEZ RODRIGUEZ JORGE EDUARDO", "CARTA SOLICITUD RESERVA DE DINERO ")
                     };
                    int index = 0;
                    foreach (var (B, C, D) in valores)
                    {
                        worksheet.Cell(filaBase, "B").Value = B;
                        worksheet.Cell(filaBase, "C").Value = C;
                        worksheet.Cell(filaBase, "D").Value = D;

                        var range = worksheet.Range($"B{filaBase}:D{filaBase}");
                        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        range.Style.Border.OutsideBorderColor = XLColor.Black;
                        range.Style.Border.InsideBorderColor = XLColor.Black;

                        if (index == 0) // Primera fila (jon, can, sal)
                        {
                            range.Style.Fill.BackgroundColor = XLColor.FromHtml("#b0c4de"); // Fondo azul claro
                            range.Style.Font.Bold = true; // Negrita
                        }

                        filaBase++;
                        index++;
                    }
                }
                }


            #endregion
            #region Auto‑ajuste de ancho de columnas
            // Ajusta el ancho de TODAS las columnas que tienen contenido
            worksheet.ColumnsUsed().AdjustToContents();
            worksheet.Column("G").Width= 15;
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