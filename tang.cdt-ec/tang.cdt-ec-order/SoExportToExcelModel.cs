using System;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;

namespace tang.cdt_ec_order
{
    public class SoExportToExcelModel
    {
        public string OrderNo { get; set; }

        public string OtherOrderId { get; set; }

        public string OrgName { get; set; }

        public string OrderTime { get; set; }

        public string SaleOrderDetailInfo { get; set; }

        public string TotalMoney { get; set; }

        public string OrderStatusName { get; set; }

        public string SkuCode { get; set; }

        public string SkuName { get; set; }

        public decimal Quantity { get; set; }

        public string Unit { get; set; }

        public decimal Price { get; set; }

        public decimal TaxPrice { get; set; }

        public decimal Deliveryleftnum { get; set; }

        public decimal DeliveredNum { get; set; }

        public string DeliverAddress { get; set; }

        public string PersonName { get; set; }

        public string Mobile { get; set; }

        public bool IsUploadFile { get; set; }

        public string IsUploadFileDisplayNote { get; set; }

        public static List<SoExportToExcelModel> Convert(SoDetailQueryResult detailResult)
        {
            List<SoExportToExcelModel> exportModels = new List<SoExportToExcelModel>();

            foreach (var saleOrderDetailRow in detailResult.DataTables.SaleOrderDetailDataTable.Rows)
            {
                string consigneeInfoStr = detailResult.DataTables.SaleOrderDataTable.Rows.First()?.Data.ConsigneeInfo;

                ConsigneeInfo consigneeInfo = JsonConvert.DeserializeObject<ConsigneeInfo>(consigneeInfoStr);

                SoExportToExcelModel exportModel = new SoExportToExcelModel
                {
                    OrderNo = detailResult.HeadSoData.OrderNo,
                    OtherOrderId = detailResult.DataTables.SaleOrderDataTable.Rows.First()?.Data.OrderOtherId,
                    OrgName = detailResult.HeadSoData.OrgName,
                    OrderTime = Util.ConvertToDateTime(detailResult.HeadSoData.OrderTime).ToString("yyyy-MM-dd"),
                    SaleOrderDetailInfo = detailResult.HeadSoData.SaleOrderDetailInfo,
                    TotalMoney = detailResult.HeadSoData.TotalMoney,
                    OrderStatusName = detailResult.HeadSoData.OrderStatusName,
                    SkuCode = saleOrderDetailRow.Data.SkuValue,
                    SkuName = saleOrderDetailRow.Data.ProductName,
                    Quantity = saleOrderDetailRow.Data.Quantity,
                    Unit = saleOrderDetailRow.Data.Unit,
                    Price = saleOrderDetailRow.Data.Price,
                    TaxPrice = saleOrderDetailRow.Data.TaxPrice,
                    Deliveryleftnum = saleOrderDetailRow.Data.Deliveryleftnum,
                    DeliveredNum = saleOrderDetailRow.Data.DeliveredNum,
                    DeliverAddress = saleOrderDetailRow.Data.DeliverAddress,
                    PersonName = consigneeInfo.NyPersonName,
                    Mobile = consigneeInfo.Mobile,
                    IsUploadFile = detailResult.IsUploadFile,
                    IsUploadFileDisplayNote = detailResult.IsUploadFileDisplayNote
                };

                exportModels.Add(exportModel);
            }

            return exportModels;
        }
    }
}

