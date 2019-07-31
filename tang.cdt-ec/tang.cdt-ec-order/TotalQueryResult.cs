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

        public HeadDeliveryOrderDataTable DeliveryOrderDataTable { get; set; }
    }

    public class HeadCustom
    {
        /// <summary>
        /// 发货管理总数量
        /// </summary>
        public int AllCount { get; set; }

        /// <summary>
        /// 待确认订单数量
        /// </summary>
        public int UnconfirmBillCount { get; set; }

        /// <summary>
        /// 已确认订单数量
        /// </summary>
        public int ConfirmedBillCount { get; set; }
    }

    public class HeadSaleOrderDataTable
    {
        public List<HeadSoRow> Rows;
    }

    public class HeadDeliveryOrderDataTable
    {
        public List<HeadDoRow> Rows;
    }

    public class HeadSoRow
    {
        public HeadSoData Data { get; set; }
    }

    public class HeadDoRow
    {
        public HeadDoData Data { get; set; }
    }

    public class HeadSoData
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

    public class HeadDoData
    {
        public int Id { get; set; }

        public string DeliveryOrderCode { get; set; }

        public string DeliveryDate { get; set; }

        public string PurchaseOrgName { get; set; }

        public string DeliverDesc { get; set; }

        public string StatusName { get; set; }
    }
}
