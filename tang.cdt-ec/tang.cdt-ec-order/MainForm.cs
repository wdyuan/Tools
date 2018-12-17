using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace tang.cdt_ec_order
{
    public partial class MainForm : Form
    {
        private const string PostUrl = @"http://tang.cdt-ec.com/import-batch/evt/dispatch";

        private const string IsUploadFileUrl = @"http://tang.cdt-ec.com/ny-mall-business/order/isUploadOrderFile?code=";

        private readonly List<DetailQueryResult> _detailResults = new List<DetailQueryResult>();

        private int _count;
        private int _totalCount;

        private delegate void DelegatePrintLogOut(string log);

        public string Cookie { get; set; }

        public MainForm()
        {
            InitializeComponent();

            ResultTextBox.ReadOnly = true;
            CheckForIllegalCrossThreadCalls = false;

            EnsureLogin();
        }

        public void EnsureLogin()
        {
            LoginForm loginForm = new LoginForm { StartPosition = FormStartPosition.CenterParent };

            loginForm.ShowDialog(this);
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            btnLoadData.Enabled = false;

            try
            {
                Task.Factory.StartNew(DataToExcel);
            }
            catch (Exception exception)
            {
                MessageBox.Show("获取数据出错，请联系管理员");

                PrintLog($@"异常信息{exception.Message}" + "\n");
            }
        }

        private void DataToExcel()
        {
            string initPostData = @"ctrl=portal.SaleOrderController&method=loadData&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderDesc%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22OrderStatusName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22erpProductVersion%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22saleOrderDetailInfo%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22ts%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.order.PurOrder%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A0%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22statuses%22%3A%5B%223%22%2C%2212%22%2C%224%22%2C%225%22%2C%226%22%2C%227%22%2C%228%22%2C%229%22%2C%2213%22%2C%2222%22%5D%2C%22pageIndex%22%3A0%2C%22pageSize%22%3A100%2C%22queryData%22%3A%7B%7D%7D";

            string initJsonResult = HttpRequestHelper.Post(PostUrl, initPostData, Cookie, Encoding.UTF8);

            TotalQueryResult initQueryResult = JsonConvert.DeserializeObject<TotalQueryResult>(initJsonResult);

            InitializeData(initQueryResult);

            HeadCustom headCustom = JsonConvert.DeserializeObject<HeadCustom>(initQueryResult.Custom);

            _totalCount = headCustom.UnconfirmBillCount;

            if (_totalCount > 100)
            {
                int loopCount = _totalCount % 100 == 0 ? _totalCount / 100 : _totalCount / 100 + 1;

                for (int i = 1; i < loopCount; i++)
                {
                    string postData = $@"ctrl=portal.SaleOrderController&method=loadData&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderDesc%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22OrderStatusName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22erpProductVersion%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22saleOrderDetailInfo%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22ts%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.order.PurOrder%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A{i}%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22statuses%22%3A%5B%223%22%2C%2212%22%2C%224%22%2C%225%22%2C%226%22%2C%227%22%2C%228%22%2C%229%22%2C%2213%22%2C%2222%22%5D%2C%22pageIndex%22%3A{i}%2C%22pageSize%22%3A100%2C%22queryData%22%3A%7B%7D%7D";

                    var jsonResult = HttpRequestHelper.Post(PostUrl, postData, Cookie, Encoding.UTF8);

                    InitializeData(jsonResult);
                }
            }

            PrintLog($@"总数据：{_totalCount} 条" + "\n");

            PrintLog(@"正在生成Excel......" + "\n");

            ConvertToExcel();
        }

        private void InitializeData(string jsonStr)
        {
            TotalQueryResult queryResult = JsonConvert.DeserializeObject<TotalQueryResult>(jsonStr);

            foreach (var headRow in queryResult.DataTables.SaleOrderDataTable.Rows)
            {
                if (headRow.Data == null) continue;

                string postData = $@"ctrl=portal.SaleOrderController&method=nyFindSaleOrderDetailById&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purchaseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22notaxMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmTotalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22purchasePhone%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplyPhone%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplyPersionName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22sendErpMsg%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderSourceId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderSource%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22enterpriseId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22enterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purchaser_info%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22consignee_info%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.saleorder.SaleOrder%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A0%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%2C%22saleOrderDetailDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22productName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22amount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmAmount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22unit%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22price%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22taxPrice%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmPrice%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22quantity%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmQuantity%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22productDescribe%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22taxrate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22pricedecidetailid%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliverEnterprise%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliverAddress%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22recvstor%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22balanceEnterprise%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22paymentEnterprise%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22buyofferdetailid%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22planDeliverDate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22confirmArriveDate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22orderDetailId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22skuId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22skuValue%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22suppProductUrl%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22customerActualReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22customerAcceptReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22diffCustomerActualReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22diffCustomerAcceptReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveredNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryleftnum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22nyDatangMaterialCode%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.saleorder.SaleOrderDetail%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A0%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22id%22%3A%22{headRow.Data.Id}%22%7D";

                string jsonResult = HttpRequestHelper.Post(PostUrl, postData, Cookie, Encoding.UTF8);

                DetailQueryResult detailResult = JsonConvert.DeserializeObject<DetailQueryResult>(jsonResult.Replace("_", ""));

                _detailResults.Add(detailResult);

                Thread.Sleep(500);

                _count++;

                PrintLog($@"正在加载，第：{_count} 条" + "\n");
            }
        }

        private void InitializeData(TotalQueryResult queryResult)
        {
            foreach (var headRow in queryResult.DataTables.SaleOrderDataTable.Rows)
            {
                if (headRow.Data == null) continue;

                string postData = $@"ctrl=portal.SaleOrderController&method=nyFindSaleOrderDetailById&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purchaseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22notaxMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmTotalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22purchasePhone%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplyPhone%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplyPersionName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22sendErpMsg%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderSourceId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderSource%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22enterpriseId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22enterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purchaser_info%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22consignee_info%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.saleorder.SaleOrder%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A0%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%2C%22saleOrderDetailDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22productName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22amount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmAmount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22unit%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22price%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22taxPrice%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmPrice%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22quantity%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22confirmQuantity%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22productDescribe%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22taxrate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22pricedecidetailid%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliverEnterprise%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliverAddress%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22recvstor%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22balanceEnterprise%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22paymentEnterprise%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22buyofferdetailid%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22planDeliverDate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22confirmArriveDate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22orderDetailId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22skuId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22skuValue%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22suppProductUrl%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22customerActualReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22customerAcceptReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22diffCustomerActualReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22diffCustomerAcceptReceivedNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveredNum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryleftnum%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22nyDatangMaterialCode%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.saleorder.SaleOrderDetail%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A0%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22id%22%3A%22{headRow.Data.Id}%22%7D";

                string jsonResult = HttpRequestHelper.Post(PostUrl, postData, Cookie, Encoding.UTF8);

                DetailQueryResult detailResult = JsonConvert.DeserializeObject<DetailQueryResult>(jsonResult.Replace("_", ""));

                string orderId = detailResult.DataTables.SaleOrderDataTable.Rows.First()?.Data.OrderOtherId;

                if (!string.IsNullOrWhiteSpace(orderId))
                {
                    string isUploadResult = HttpRequestHelper.Get(IsUploadFileUrl + orderId, Cookie, Encoding.UTF8);

                    IsUploadFileResult isUploadFileResult = JsonConvert.DeserializeObject<IsUploadFileResult>(isUploadResult);

                    if (isUploadFileResult.Result.Equals("fail", StringComparison.OrdinalIgnoreCase))
                    {
                        detailResult.IsUploadFile = "否";
                    }
                }

                _detailResults.Add(detailResult);

                _count++;

                PrintLog($@"正在加载，第：{_count} 条" + "\n");
            }
        }

        private void ConvertToExcel()
        {
            List<ExportToExcelModel> allExportData = new List<ExportToExcelModel>();

            foreach (var detailResult in _detailResults)
            {
                allExportData.AddRange(ExportToExcelModel.Convert(detailResult));
            }

            IWorkbook workbook = new HSSFWorkbook();

            ISheet sheet = workbook.CreateSheet("大唐电子");
            IRow row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue("采购订单号");
            row0.CreateCell(1).SetCellValue("商品编码");
            row0.CreateCell(2).SetCellValue("商品名称");
            row0.CreateCell(3).SetCellValue("数量");
            row0.CreateCell(4).SetCellValue("单位");
            row0.CreateCell(5).SetCellValue("无税单价");
            row0.CreateCell(6).SetCellValue("含税单价");
            row0.CreateCell(7).SetCellValue("可发货数量");
            row0.CreateCell(8).SetCellValue("已发货数量");
            row0.CreateCell(9).SetCellValue("收货地址");
            row0.CreateCell(10).SetCellValue("收货人");
            row0.CreateCell(11).SetCellValue("联系方式");
            row0.CreateCell(12).SetCellValue("是否上传附件");

            int index = 1;

            foreach (var exportModel in allExportData)
            {
                IRow row = sheet.CreateRow(index);

                row.CreateCell(0).SetCellValue(exportModel.OrderId);
                row.CreateCell(1).SetCellValue(exportModel.SkuCode);
                row.CreateCell(2).SetCellValue(exportModel.SkuName);
                row.CreateCell(3).SetCellValue(exportModel.Quantity.ToString("N0"));
                row.CreateCell(4).SetCellValue(exportModel.Unit);
                row.CreateCell(5).SetCellValue(exportModel.Price.ToString("N2"));
                row.CreateCell(6).SetCellValue(exportModel.TaxPrice.ToString("N2"));
                row.CreateCell(7).SetCellValue(exportModel.Deliveryleftnum.ToString("N0"));
                row.CreateCell(8).SetCellValue(exportModel.DeliveredNum.ToString("N0"));
                row.CreateCell(9).SetCellValue(exportModel.DeliverAddress);
                row.CreateCell(10).SetCellValue(exportModel.PersonName);
                row.CreateCell(11).SetCellValue(exportModel.Mobile);
                row.CreateCell(12).SetCellValue(exportModel.IsUploadFile);

                index++;
            }

            string dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            string exportFileName = $@"{dir}\大唐电子商务订单信息{DateTime.Now:yyyyMMdd}.xls";

            using (FileStream url = File.OpenWrite(exportFileName))
            {
                workbook.Write(url);
            }

            PrintLog($@"实际生成数据：{_detailResults.Count} 条" + "\n");

            PrintLog("Excel生成完成");

            btnLoadData.Enabled = true;
        }

        private void PrintLog(string log)
        {
            if (ResultTextBox.InvokeRequired)
            {
                DelegatePrintLogOut printLogOut = PrintLog;

                Invoke(printLogOut, log);
            }
            else
            {
                ResultTextBox.AppendText(log);
            }
        }
    }
}
