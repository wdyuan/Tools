using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace tang.cdt_ec_order
{
    public class DoExportToExcelModel
    {
        public string DeliveryOrderCode { get; set; }

        public string DeliveryDate { get; set; }

        public string PurchaseOrgName { get; set; }

        public string StatusName { get; set; }

        public string SaleOrderCode { get; set; }

        public string ProductSubject { get; set; }

        public decimal Num { get; set; }

        public string Unit { get; set; }

        public string PlanArrivalDate { get; set; }

        public string ReceiveAddr { get; set; }

        public string NyPersonName { get; set; }

        public string Mobile { get; set; }

        public string Bmemo { get; set; }

        public static List<DoExportToExcelModel> Convert(DoDetailQueryResult detailResult)
        {
            List<DoExportToExcelModel> exportModels = new List<DoExportToExcelModel>();

            foreach (var doDetailRow in detailResult.Data.Details)
            {
                string consigneeInfoStr = detailResult.Data.Details.First()?.ConsigneeInfo;

                ConsigneeInfo consigneeInfo = JsonConvert.DeserializeObject<ConsigneeInfo>(consigneeInfoStr);

                DoExportToExcelModel exportModel = new DoExportToExcelModel
                {
                    DeliveryOrderCode = detailResult.HeadDoData.DeliveryOrderCode,
                    DeliveryDate = string.IsNullOrWhiteSpace(detailResult.HeadDoData.DeliveryDate) ? string.Empty : Util.ConvertToDateTime(detailResult.HeadDoData.DeliveryDate).ToString("yyyy-MM-dd"),
                    PurchaseOrgName = detailResult.HeadDoData.PurchaseOrgName,
                    StatusName = detailResult.HeadDoData.StatusName,
                    SaleOrderCode = doDetailRow.SaleOrderCode,
                    ProductSubject = doDetailRow.ProductSubject,
                    Num = doDetailRow.Num,
                    Unit = doDetailRow.Unit,
                    PlanArrivalDate = string.IsNullOrWhiteSpace(doDetailRow.PlanArrivalDate) ? string.Empty : Util.ConvertToDateTime(doDetailRow.PlanArrivalDate).ToString("yyyy-MM-dd"),
                    ReceiveAddr = doDetailRow.ReceiveAddr,
                    NyPersonName = consigneeInfo.NyPersonName,
                    Mobile = consigneeInfo.Mobile,
                    Bmemo = doDetailRow.Bmemo
                };

                exportModels.Add(exportModel);
            }

            return exportModels;
        }
    }
}
