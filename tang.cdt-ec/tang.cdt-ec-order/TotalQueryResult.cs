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
    }
}
