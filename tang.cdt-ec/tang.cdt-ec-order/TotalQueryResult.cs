using System.Collections.Generic;

namespace tang.cdt_ec_order
{
    public class TotalQueryResult
    {
        public HeadDataTable DataTables { get; set; }

        public string Custom { get; set; }
    }

    public class HeadDataTable
    {
        public HeadSaleOrderDataTable SaleOrderDataTable { get; set; }
    }

    public class HeadCustom
    {
        public int UnconfirmBillCount { get; set; }

        public int ConfirmedBillCount { get; set; }
    }

    public class HeadSaleOrderDataTable
    {
        public List<HeadRow> Rows;
    }

    public class HeadRow
    {
        public HeadData Data { get; set; }
    }

    public class HeadData
    {
        public int Id { get; set; }

        public string OrderNo { get; set; }

        public string OrderOtherId { get; set; }

        public string OrgName { get; set; }

        public string OrderTime { get; set; }

        public string SaleOrderDetailInfo { get; set; }

        public string TotalMoney { get; set; }

        public string OrderStatusName { get; set; }
    }
}
