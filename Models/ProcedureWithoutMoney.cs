using System.Diagnostics.CodeAnalysis;

namespace WebApplication1.Models
{
    public class ProcedureWithoutMoney : BaseEntity
    {
        public required string InsuranceCode { get; set; } // poliza
        public required string Insured { get; set; } // nombre
        public required string Procedure { get; set; } // tramite

        private ProcedureWithoutMoney() { }

        [SetsRequiredMembers]
        public ProcedureWithoutMoney(string insuranceCode, string insured, string procedure)
        {
            InsuranceCode = insuranceCode;
            Insured = insured;
            Procedure = procedure;
        }
    }
}
