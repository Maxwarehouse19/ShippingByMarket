using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShippingByMarket
{
    class PedidoEndicia
    {
        public DateTime PrintDate { get; set; }
        public Decimal AmountPaid { get; set; }
        public string AdjAmount { get; set; }
        public Decimal QuotedAmount { get; set; }
        public string Recipient { get; set; }
        public string Status { get; set; }
        public string TrackingNumber { get; set; }
        public DateTime DateDelivered { get; set; }
        public string Carrier { get; set; }
        public string ClassService { get; set; }
        public Decimal InsuredValue { get; set; }
        public string InsuranceID { get; set; }
        public string CostCode { get; set; }
        public string Weight { get; set; }
        public DateTime ShipDate { get; set; }
        public string RefundType { get; set; }
        public string PrintedMessage { get; set; }
        public string User { get; set; }
        public DateTime RefundRequestDate { get; set; }
        public string RefundStatus { get; set; }
        public string RefundRequested { get; set; }
        public string Reference1 { get; set; }
        public string Reference2 { get; set; }
        public string Reference3 { get; set; }
        public string Reference4 { get; set; }
    }

}
