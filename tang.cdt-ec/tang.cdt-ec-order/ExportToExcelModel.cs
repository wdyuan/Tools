using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;

namespace tang.cdt_ec_order
{
    public class ExportToExcelModel
    {
        public string OrderId { get; set; }

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

        public static List<ExportToExcelModel> Convert(DetailQueryResult detailResult)
        {
            List<ExportToExcelModel> exportModels = new List<ExportToExcelModel>();

            foreach (var saleOrderDetailRow in detailResult.DataTables.SaleOrderDetailDataTable.Rows)
            {
                string consigneeInfoStr = detailResult.DataTables.SaleOrderDataTable.Rows.First()?.Data.ConsigneeInfo;

                ConsigneeInfo consigneeInfo = JsonConvert.DeserializeObject<ConsigneeInfo>(consigneeInfoStr);

                ExportToExcelModel exportModel = new ExportToExcelModel
                {
                    OrderId = detailResult.DataTables.SaleOrderDataTable.Rows.First()?.Data.OrderOtherId,
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
                    Mobile = consigneeInfo.Mobile
                };

                exportModels.Add(exportModel);
            }

            return exportModels;
        }
    }
}

