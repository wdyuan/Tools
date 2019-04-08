using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace tang.cdt_ec_order
{
    public partial class MainForm : Form
    {
        private const string PostUrl = @"http://tang.cdt-ec.com/import-batch/evt/dispatch";

        private const string IsUploadFileUrl = @"http://tang.cdt-ec.com/ny-mall-business/order/isUploadOrderFile?code=";

        private const string DownloadFileUrl = @"http://tang.cdt-ec.com/ny-mall-business/order/downLoadOrderFileByCode?code=";

        private readonly Dictionary<SoType, List<DetailQueryResult>> _typeDetails = new Dictionary<SoType, List<DetailQueryResult>>
        {
            {SoType.Confirmed, new List<DetailQueryResult>()},
            {SoType.UnConfirm, new List<DetailQueryResult>()}
        };

        private int _unConfirmCount;
        private int _confirmedCount;
        private int _totalUnConfirmCount;
        private int _totalConfirmedCount;

        private delegate void DelegatePrintLogOut(string log);

        public string Cookie { get; set; }

        public enum SoType
        {
            Confirmed,
            UnConfirm
        }

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
            string unConfirmPostData = @"ctrl=portal.SaleOrderController&method=loadData&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderDesc%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22OrderStatusName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22erpProductVersion%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22saleOrderDetailInfo%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22ts%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.order.PurOrder%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A0%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22statuses%22%3A%5B%223%22%2C%2212%22%2C%224%22%2C%225%22%2C%226%22%2C%227%22%2C%228%22%2C%229%22%2C%2213%22%2C%2222%22%5D%2C%22pageIndex%22%3A0%2C%22pageSize%22%3A100%2C%22queryData%22%3A%7B%7D%7D";

            string confirmedPostData = @"ctrl=portal.SaleOrderController&method=loadData&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderDesc%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22OrderStatusName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22erpProductVersion%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22saleOrderDetailInfo%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22ts%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.order.PurOrder%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A0%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22statuses%22%3A%5B%2210%22%2C%2211%22%2C%2214%22%2C%2215%22%2C%2216%22%2C%2218%22%2C%2219%22%5D%2C%22pageIndex%22%3A0%2C%22pageSize%22%3A10%2C%22queryData%22%3A%7B%7D%7D";

            QueryToExportData(unConfirmPostData, SoType.UnConfirm);

            QueryToExportData(confirmedPostData, SoType.Confirmed);

            PrintLog(@"正在生成Excel......" + "\n");

            ConvertToExcel();
        }

        private void QueryToExportData(string postData, SoType soType)
        {
            int loopCount = 1;

            string queryResult = HttpRequestHelper.Post(PostUrl, postData, Cookie, Encoding.UTF8);

            TotalQueryResult initQueryResult = JsonConvert.DeserializeObject<TotalQueryResult>(queryResult);

            HeadCustom headCustom = JsonConvert.DeserializeObject<HeadCustom>(initQueryResult.Custom);

            switch (soType)
            {
                case SoType.UnConfirm:
                    {
                        _totalUnConfirmCount += headCustom.UnconfirmBillCount;

                        PrintLog($@"未确认数据：{_totalUnConfirmCount} 条" + "\n");

                        if (_totalUnConfirmCount <= 100) return;

                        loopCount = _totalUnConfirmCount % 100 == 0 ? _totalUnConfirmCount / 100 : _totalUnConfirmCount / 100 + 1;

                        break;
                    }
                case SoType.Confirmed:
                    {
                        _totalConfirmedCount += headCustom.ConfirmedBillCount;

                        PrintLog($@"已确认数据：{_totalConfirmedCount} 条" + "\n");

                        if (_totalConfirmedCount <= 100) return;

                        loopCount = _totalConfirmedCount % 100 == 0 ? _totalConfirmedCount / 100 : _totalConfirmedCount / 100 + 1;

                        break;
                    }
            }

            for (int i = 0; i < loopCount; i++)
            {
                string paginationPostData = soType == SoType.UnConfirm
                    ? $@"ctrl=portal.SaleOrderController&method=loadData&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderDesc%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22OrderStatusName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22erpProductVersion%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22saleOrderDetailInfo%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22ts%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.order.PurOrder%22%7D%2C%22rows%22%3A%5B%5D%2C%22select%22%3A%5B%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A10%2C%22pageIndex%22%3A{i}%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22statuses%22%3A%5B%223%22%2C%2212%22%2C%224%22%2C%225%22%2C%226%22%2C%227%22%2C%228%22%2C%229%22%2C%2213%22%2C%2222%22%5D%2C%22pageIndex%22%3A{i}%2C%22pageSize%22%3A100%2C%22queryData%22%3A%7B%7D%7D"
                    : $@"ctrl=portal.SaleOrderController&method=loadData&environment=%7B%22clientAttributes%22%3A%7B%7D%7D&dataTables=%7B%22saleOrderDataTable%22%3A%7B%22meta%22%3A%7B%22id%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22subject%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderDesc%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderno%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderOtherId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderTime%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22date%22%7D%2C%22corpAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22corpSubAccount%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22supplierName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22totalMoney%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orderStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22OrderStatusName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22gmtCreate%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseStart%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22releaseEnd%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%2C%22type%22%3A%22datetime%22%7D%2C%22supEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22erpProductVersion%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22saleOrderDetailInfo%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purEnterpriseName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22deliveryStatus%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22orgName%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22purOrderId%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%2C%22ts%22%3A%7B%22enable%22%3Atrue%2C%22required%22%3Afalse%2C%22descs%22%3A%7B%7D%7D%7D%2C%22params%22%3A%7B%22cls%22%3A%22com.yonyou.cpu.domain.order.PurOrder%22%7D%2C%22rows%22%3A%5B%7B%22id%22%3A%22r95638%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%22id%22%3A%7B%22value%22%3A29273%7D%2C%22subject%22%3A%7B%22value%22%3Anull%7D%2C%22orderDesc%22%3A%7B%22value%22%3Anull%7D%2C%22orderno%22%3A%7B%22value%22%3A%22O-CX-19-00006599-1%22%7D%2C%22orderOtherId%22%3A%7B%22value%22%3A%22O-CX-19-00006599%22%7D%2C%22orderTime%22%3A%7B%22value%22%3A1554652800000%7D%2C%22corpAccount%22%3A%7B%22value%22%3Anull%7D%2C%22corpSubAccount%22%3A%7B%22value%22%3Anull%7D%2C%22supplierName%22%3A%7B%22value%22%3Anull%7D%2C%22totalMoney%22%3A%7B%22value%22%3A286.52%7D%2C%22orderStatus%22%3A%7B%22value%22%3A%2210%22%7D%2C%22OrderStatusName%22%3A%7B%22value%22%3A%22%E5%BE%85%E5%8F%91%E8%B4%A7%22%7D%2C%22gmtCreate%22%3A%7B%22value%22%3A1554688601000%7D%2C%22releaseStart%22%3A%7B%22value%22%3Anull%7D%2C%22releaseEnd%22%3A%7B%22value%22%3Anull%7D%2C%22supEnterpriseName%22%3A%7B%22value%22%3Anull%7D%2C%22erpProductVersion%22%3A%7B%22value%22%3Anull%7D%2C%22saleOrderDetailInfo%22%3A%7B%22value%22%3A%22%E8%B5%84%E7%94%9F%E5%A0%82+SHISEIDO+%E6%B0%B4%E4%B9%8B%E5%AF%86%E8%AF%AD+%E6%B2%90%E6%B5%B4%E9%9C%B2+%E6%B5%B7%E7%9B%90%E5%BC%B9%E6%B6%A6%E7%B4%A7+600ml+%EF%BC%8C1.00(%E7%93%B6)%3B%E8%B5%84%E7%94%9F%E5%A0%82+SHISEIDO+%E6%B0%B4%E4%B9%8B%E5%AF%86%E8%AF%AD%E5%87%80%E6%BE%84%E6%B0%B4%E6%B4%BB%E6%8A%A4%E5%8F%91%E7%B4%A0+600ml%2F%E7%93%B6+9%E7%93%B6%2F%E7%AE%B1+%EF%BC%8C1.00(%E7%93%B6)%3B%E8%B5%84%E7%94%9F%E5%A0%82+SHISEIDO+%E7%BE%8E%E6%B6%A6%E6%8A%A4%E6%89%8B%E9%9C%9C+100g%2F%E7%9B%92+(%E6%B8%97%E9%80%8F%E6%BB%8B%E5%85%BB%E5%9E%8B)+%EF%BC%8C2.00(%E7%9B%92)%3B%E8%B5%84%E7%94%9F%E5%A0%82+SHISEIDO+%E6%B0%B4%E4%B9%8B%E5%AF%86%E8%AF%AD%E5%87%80%E6%BE%84%E6%B0%B4%E6%B4%BB%E6%B4%97%E5%8F%91%E9%9C%B2+600ml%2F%E7%93%B6+9%E7%93%B6%2F%E7%AE%B1+(%E5%80%8D%E6%B6%A6%E5%9E%8B)+%EF%BC%8C1.00(%E7%93%B6)%3B%22%7D%2C%22purEnterpriseName%22%3A%7B%22value%22%3A%22%E5%A4%A7%E5%94%90%E7%94%B5%E5%95%86%22%7D%2C%22deliveryStatus%22%3A%7B%22value%22%3A%221%22%7D%2C%22orgName%22%3A%7B%22value%22%3A%22%E4%B8%AD%E5%9B%BD%E6%B0%B4%E5%88%A9%E7%94%B5%E5%8A%9B%E7%89%A9%E8%B5%84%E5%8C%97%E4%BA%AC%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%7D%2C%22purOrderId%22%3A%7B%22value%22%3A1000018835%7D%2C%22ts%22%3A%7B%22value%22%3A1554690703000%7D%7D%7D%2C%7B%22id%22%3A%22r62760%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r34117%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r95830%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r43874%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r47452%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r39154%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r46075%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r47073%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%2C%7B%22id%22%3A%22r36335%22%2C%22status%22%3A%22nrm%22%2C%22data%22%3A%7B%7D%7D%5D%2C%22select%22%3A%5B0%5D%2C%22focus%22%3A-1%2C%22pageSize%22%3A100%2C%22pageIndex%22%3A{i}%2C%22isChanged%22%3Afalse%2C%22master%22%3A%22%22%2C%22pageCache%22%3Afalse%7D%7D&compression=false&compressType=&parameters=%7B%22statuses%22%3A%5B%2210%22%2C%2211%22%2C%2214%22%2C%2215%22%2C%2216%22%2C%2218%22%2C%2219%22%5D%2C%22pageIndex%22%3A{i}%2C%22pageSize%22%3A100%2C%22queryData%22%3A%7B%7D%7D";

                var jsonResult = HttpRequestHelper.Post(PostUrl, paginationPostData, Cookie, Encoding.UTF8);

                TotalQueryResult paginationqueryResult = JsonConvert.DeserializeObject<TotalQueryResult>(jsonResult);

                InitializeDetail(soType, paginationqueryResult);
            }
        }

        private void InitializeDetail(SoType soType, TotalQueryResult queryResult)
        {
            string typeName = soType == SoType.Confirmed ? "已确认" : "待确认";

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
                        detailResult.IsUploadFile = false;
                        detailResult.IsUploadFileDisplayNote = "否";
                    }
                    else
                    {
                        detailResult.IsUploadFile = true;
                        detailResult.IsUploadFileDisplayNote = "是";
                    }
                }

                switch (soType)
                {
                    case SoType.UnConfirm:
                        {
                            _typeDetails[SoType.UnConfirm].Add(detailResult);

                            _unConfirmCount++;

                            PrintLog($@"类型：{typeName}，正在加载，第：{_unConfirmCount} 条" + "\n");

                            break;
                        }
                    case SoType.Confirmed:
                        {
                            _typeDetails[SoType.Confirmed].Add(detailResult);

                            _confirmedCount++;

                            PrintLog($@"类型：{typeName}，正在加载，第：{_confirmedCount} 条" + "\n");

                            break;
                        }
                }
            }
        }

        private void ConvertToExcel()
        {
            IWorkbook workbook = new HSSFWorkbook();

            ISheet unConfirmSheet = workbook.CreateSheet("未确认");

            ISheet confirmedSheet = workbook.CreateSheet("已确认");

            List<ExportToExcelModel> unConfirmModels = _typeDetails[SoType.UnConfirm].SelectMany(ExportToExcelModel.Convert).ToList();

            List<ExportToExcelModel> confirmedModels = _typeDetails[SoType.Confirmed].SelectMany(ExportToExcelModel.Convert).ToList();

            FillSheetData(unConfirmSheet, unConfirmModels);

            FillSheetData(confirmedSheet, confirmedModels);

            string dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            string exportFileName = $@"{dir}\大唐电子商务订单信息{DateTime.Now:yyyyMMdd}.xls";

            using (FileStream url = File.OpenWrite(exportFileName))
            {
                workbook.Write(url);
            }

            string uploadFilePath = $@"{dir}\附件信息";

            if (!Directory.Exists(uploadFilePath))
            {
                Directory.CreateDirectory(uploadFilePath);
            }

            foreach (var groupModel in unConfirmModels.Where(model => model.IsUploadFile).GroupBy(model => model.OrderId))
            {
                string orderId = groupModel.Key;

                string fileName = $@"未处理-附件{orderId}.xls";

                HttpRequestHelper.HttpDownloadFile(DownloadFileUrl + orderId, uploadFilePath, fileName, true);
            }

            //foreach (var groupModel in confirmedModels.Where(model => model.IsUploadFile).GroupBy(model => model.OrderId))
            //{
            //    string orderId = groupModel.Key;

            //    string fileName = $@"已处理-附件{orderId}.xls";

            //    HttpRequestHelper.HttpDownloadFile(DownloadFileUrl + orderId, uploadFilePath, fileName, true);
            //}

            PrintLog($@"未确认数据，实际生成：{_unConfirmCount} 条" + "\n");

            PrintLog($@"已确认数据，实际生成：{_confirmedCount} 条" + "\n");

            PrintLog("Excel生成完成");

            btnLoadData.Enabled = true;
        }

        /// <summary>
        /// 填充Sheet数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="exportToExcelModels"></param>
        private void FillSheetData(ISheet sheet, List<ExportToExcelModel> exportToExcelModels)
        {
            IRow headRow = sheet.CreateRow(0);

            headRow.CreateCell(0).SetCellValue("采购订单号");
            headRow.CreateCell(1).SetCellValue("商品编码");
            headRow.CreateCell(2).SetCellValue("商品名称");
            headRow.CreateCell(3).SetCellValue("数量");
            headRow.CreateCell(4).SetCellValue("单位");
            headRow.CreateCell(5).SetCellValue("无税单价");
            headRow.CreateCell(6).SetCellValue("含税单价");
            headRow.CreateCell(7).SetCellValue("可发货数量");
            headRow.CreateCell(8).SetCellValue("已发货数量");
            headRow.CreateCell(9).SetCellValue("收货地址");
            headRow.CreateCell(10).SetCellValue("收货人");
            headRow.CreateCell(11).SetCellValue("联系方式");
            headRow.CreateCell(12).SetCellValue("是否上传附件");

            int index = 1;

            foreach (var exportModel in exportToExcelModels)
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
                row.CreateCell(12).SetCellValue(exportModel.IsUploadFileDisplayNote);

                index++;
            }
        }

        /// <summary>
        /// 输出日志
        /// </summary>
        /// <param name="log"></param>
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
