﻿using System.Collections.Generic;

namespace tang.cdt_ec_order
{
    public class SoDetailQueryResult
    {
        public HeadSoData HeadSoData { get; set; }

        public bool IsUploadFile { get; set; }

        public string IsUploadFileDisplayNote { get; set; }

        public SoDetailDataTable DataTables { get; set; }
    }

    public class SoDetailDataTable
    {
        public SaleOrderDataTable SaleOrderDataTable { get; set; }

        public SaleOrderDetailDataTable SaleOrderDetailDataTable { get; set; }
    }

    public class SaleOrderDataTable
    {
        public List<SaleOrderRow> Rows { get; set; }
    }

    public class SaleOrderDetailDataTable
    {
        public List<SaleOrderDetailRow> Rows { get; set; }
    }

    public class SaleOrderRow
    {
        public SaleOrderData Data { get; set; }
    }

    public class SaleOrderDetailRow
    {
        public SaleOrderDetailData Data { get; set; }
    }

    public class SaleOrderData
    {
        public string OrderOtherId { get; set; }

        public string ConsigneeInfo { get; set; }
    }

    public class SaleOrderDetailData
    {
        public string SkuValue { get; set; }

        public string ProductName { get; set; }

        public decimal Quantity { get; set; }

        public string Unit { get; set; }

        public decimal TaxPrice { get; set; }

        public decimal Price { get; set; }

        public decimal Deliveryleftnum { get; set; }

        public decimal DeliveredNum { get; set; }

        public string DeliverAddress { get; set; }
    }

    public class ConsigneeInfo
    {
        public string NyPersonName { get; set; }

        public string Mobile { get; set; }
    }

    public class IsUploadFileResult
    {
        public string Msg { get; set; }

        public string Result { get; set; }
    }
}
