using System.Diagnostics.CodeAnalysis;

namespace WebApplication1.Models
{
    public class DailyLog : BaseEntity
    {
        public required int InsuranceId { get; set; }
        public required string Code { get; set; } // Numero Poliza
        public required string ReceiptNumber { get; set; } // Numero Rec/Com
        public required string InsuranceType { get; set; } // Tipo Seguro
        public required DateTime From { get; set; } // Vencimiento Desde
        public required DateTime To { get; set; } // Vencimiento Hasta
        public required decimal Amount { get; set; } //Monto
        public required string PaymentMethod { get; set; } //Medio de pago e.g. V-123 o EFECTIVO o D-456
        public required string TransactionType { get; set; } // Tipo de transaccion
        public required string InsuredFullName { get; set; } // asegurado
        public required string Identification { get; set; } // cedula
        public required DateTime DailyLogDate { get; set; } // fecha de bitacora
        public string? PhoneNumber { get; set; } // telefono
        public string? LicensePlate { get; set; } //placa
        public string? ChargeDay { get; set; } // cargo manual 

        private DailyLog() { }

        [SetsRequiredMembers]
        public DailyLog(
            string receiptNumber,
            string code,
            string insuranceType,
            DateTime from,
            DateTime to,
            decimal amount,
            string paymentMethod,
            string transactionType,
            string insuredFullName,
            string identification,
            DateTime dailyLogDate,
            string? phoneNumber = null,
            string? licensePlate = null,
            string? chargeDay = null)
        {
            ReceiptNumber = receiptNumber;
            Code = code;
            InsuranceType = insuranceType;
            From = from;
            To = to;
            Amount = amount;
            PaymentMethod = paymentMethod;
            TransactionType = transactionType;
            InsuredFullName = insuredFullName;
            Identification = identification;
            DailyLogDate = dailyLogDate;
            PhoneNumber = phoneNumber;
            LicensePlate = licensePlate;
            ChargeDay = chargeDay;
        }
    }
}
