using System.Collections.Generic;

namespace tang.cdt_ec_order
{
    public class DoDetailQueryResult
    {
        public HeadDoData HeadDoData { get; set; }

        public DoDetailData Data { get; set; }
    }

    public class DoDetailData
    {
        public List<DoDetailRow> Details;
    }

    public class DoDetailRow
    {
        public string SaleOrderCode { get; set; }

        public string ProductSubject { get; set; }

        public decimal Num { get; set; }

        public string Unit { get; set; }

        public string PlanArrivalDate { get; set; }

        public string ReceiveAddr { get; set; }

        public string Bmemo { get; set; }

        public string ConsigneeInfo { get; set; }
    }
}
