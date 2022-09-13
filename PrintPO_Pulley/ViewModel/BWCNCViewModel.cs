
using BDDataCrawler.Provider;
using BFDataCrawler.BLL;
using BFDataCrawler.Helper;
using BFDataCrawler.Helpers;
using BFDataCrawler.Model;
using BFDataCrawler.View;
using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Syroot.Windows.IO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Navigation;
using System.Windows.Threading;
using waitHelpers = SeleniumExtras.WaitHelpers;
namespace BFDataCrawler.ViewModel
{
    class BWCNCViewModel : INotifyPropertyChanged
    {        
        #region <<CONSTRUCTOR>>
        public BWCNCViewModel()
        {
            
            BWOrders = new ObservableCollection<OrderModel>();
            CNCOrders = new ObservableCollection<OrderModel>();

            //initialize instance of CNC
            ListOrdersCNCHob9 = new ObservableCollection<OrderModel>();
            ListOrdersCNCHob10 = new ObservableCollection<OrderModel>();
            ListOrdersCNCNoHob = new ObservableCollection<OrderModel>();
            //initialize instance of BW
            ListOrdersBWHob11 = new ObservableCollection<OrderModel>();
            ListOrdersBWHob12 = new ObservableCollection<OrderModel>();
            ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>();
            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>();
            ListOrdersBWNoHob = new ObservableCollection<OrderModel>();
            ListOrdersBWF1 = new ObservableCollection<OrderModel>();


            IsLoadingBW = false;
            IsGettingDataBW = true;

            IsLoadingCNC = false;
            IsGettingDataCNC = true;

            BWHob11_IsGettingData = true;
            BWHob11_IsPrinting = false;

            BWHob12_IsGettingData = true;
            BWHob12_IsPrinting = false;

            BWNoHob_IsGettingData = true;
            BWNoHob_IsPrinting = false;

            BWF1_IsGettingData = true;
            BWF1_IsPrinting = false;

            BWCNCHob9_IsGettingData = true;
            BWCNCHob9_IsPrinting = false;

            BWCNCHob10_IsGettingData = true;
            BWCNCHob10_IsPrinting = false;

            CNCHob9_IsGettingData = true;
            CNCHob9_IsPrinting = false;

            CNCHob10_IsGettingData = true;
            CNCHob10_IsPrinting = false;

            //2019-12-03 Start add:
            Hob9BWCNC_TeethShape_IsEnabled = false;
            Hob9BWCNC_TeethQty_IsEnabled = false;
            Hob9BWCNC_Diameter_IsEnabled = false;
            Hob9BWCNC_GlobalCode_IsEnabled = false;

            Hob10BWCNC_TeethShape_IsEnabled = false;
            Hob10BWCNC_TeethQty_IsEnabled = false;
            Hob10BWCNC_Diameter_IsEnabled = false;
            Hob10BWCNC_GlobalCode_IsEnabled = false;

            Hob9CNC_TeethShape_IsEnabled = false;
            Hob9CNC_TeethQty_IsEnabled = false;
            Hob9CNC_Diameter_IsEnabled = false;
            Hob9CNC_Global_IsEnabled = false;

            Hob10CNC_TeethShape_IsEnabled = false;
            Hob10CNC_TeethQty_IsEnabled = false;
            Hob10CNC_Diameter_IsEnabled = false;
            Hob10CNC_Global_IsEnabled = false;
            //2019-12-03 End add

            //2020-03-30 Start add: lấy teethInfo của BW11 và BW12 đồng bộ từ server về
            try
            {
                var teethInfoBW11 = OrderBLL.Instance.GetBW11_TeethInfo();
                var teethInfoBW12 = OrderBLL.Instance.GetBW12_TeethInfo();
                if (teethInfoBW11 != null)
                {
                    Hob11BW_TeethShape = teethInfoBW11["TeethShape"].ToString();
                    Hob11BW_TeethQty = Convert.ToInt32(teethInfoBW11["TeethQty"].ToString());
                    Hob11BW_Diameter = Convert.ToDouble(teethInfoBW11["Diameter_d"].ToString());
                    Hob11BW_GlobalCode = teethInfoBW11["GlobalCode"].ToString();
                }

                if(teethInfoBW12 != null)
                {
                    Hob12BW_TeethShape = teethInfoBW12["TeethShape"].ToString();
                    Hob12BW_TeethQty = Convert.ToInt32(teethInfoBW12["TeethQty"].ToString());
                    Hob12BW_Diameter = Convert.ToDouble(teethInfoBW12["Diameter_d"].ToString());
                    Hob12BW_GlobalCode = teethInfoBW12["GlobalCode"].ToString();
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Lấy thông số dao, đường kính cho máy BW thất bại, vui lòng kiểm tra lại");
            }
         
            //2019-12-05 Start add: lock, unlock điều chỉnh thông số sắp xếp orders, block che datagrid
            BW_IsUnLockSettings = false;
            CNC_IsUnLockSettings = false;

            BW_BlockIsVisible = Visibility.Visible;
            CNC_BlockIsVisible = Visibility.Visible;
            //2019-12-05 End add

            //KillChromeDriver();

            //2019-10-14 Start add: auto synchronize deleted item from ListOrderBWHob
            timerBWHob = new DispatcherTimer()
            {
                Interval = TimeSpan.FromSeconds(10)
            };
            //timerBWHob.Start();
            timerBWHob.Tick += new EventHandler((s, e) =>
            {                
                try
                {                   
                    OrderModel orderCNCHob9 = new OrderModel(MySQLProvider.Instance.ExecuteQuery("SELECT * FROM CNCHob9").Rows[0]);
                    OrderModel orderCNCHob10 = new OrderModel(MySQLProvider.Instance.ExecuteQuery("SELECT * FROM CNCHob10").Rows[0]);
                    //2019-12-02 End change
                    if (orderCNCHob9 != null && ListOrdersBWCNCHob9 != null)
                    {
                        var order = ListOrdersBWCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No.Equals(orderCNCHob9.Manufa_Instruction_No));
                        if (order != null)
                            ListOrdersBWCNCHob9.Remove(order);

                        if (orderCNCHob10 != null)
                        {
                            var order2 = ListOrdersBWCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No.Equals(orderCNCHob10.Manufa_Instruction_No));
                            if (order2 != null)
                                ListOrdersBWCNCHob9.Remove(order2);
                        }
                    }
                    //2019-10-17 End add

                    //2019-10-24 Start add: xoa PO cua Hob10 BWCNC                
                    if (orderCNCHob10 != null && ListOrdersBWCNCHob10 != null)
                    {
                        var order = ListOrdersBWCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No.Equals(orderCNCHob10.Manufa_Instruction_No));
                        if (order != null)
                            ListOrdersBWCNCHob10.Remove(order);

                        if (orderCNCHob9 != null)
                        {
                            var order2 = ListOrdersBWCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No.Equals(orderCNCHob9.Manufa_Instruction_No));
                            if (order2 != null)
                                ListOrdersBWCNCHob10.Remove(order2);
                        }
                    }
                    //2019-10-24 End add

                    //2020-02-08 Start edit: cân bằng đơn của BWCNC9 và 10, nếu vẫn còn máy hết đơn thì chia sẻ từ BW qua
                    BalancedOrdersBWCNC();
                    if (ListOrdersBWCNCHob9 == null) { ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(); }
                    if (ListOrdersBWCNCHob10 == null) { ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(); }
                    //2020-02-08 End edit

                    //2019-12-19 Start add: cân bằng số lượng orders cho máy BW
                    //2020-02-24 Start edit: chuyển sang áp dụng check NYP (SPH và PT)
                    //BalancedOrdersBW();
                    var ordersBW11 = ListOrdersBWHob11 != null ? ListOrdersBWHob11 : (new ObservableCollection<OrderModel>());
                    var ordersBW12 = ListOrdersBWHob12 != null ? ListOrdersBWHob12 : (new ObservableCollection<OrderModel>());
                    ApplySSDCheck_BW(ordersBW11.Union(ordersBW12).Union(ListOrdersBWCNCHob9.Where(c=>c.LineAB.Equals("BW"))).Union(ListOrdersBWCNCHob10.Where(c=>c.LineAB.Equals("BW"))).ToList());

                    int status = OrderBLL.Instance.Get_bwIsSetupN1();
                    if(status == 1) //cân bằng có đơn N+1
                    {
                        if(ListOrdersBWCNCHob9.Count == 0 || ListOrdersBWCNCHob10.Count ==0)
                        {
                            SharingBWOrdersToCNC_SSDMode();//tạm thời chưa làm để test GetOrdersBW_N1
                            //SharingBWOrdersToCNC_SSDMode();
                        }
                        BalancedOrdersBW("CheckSSD");
                    }
                    else
                    {                        
                        if (ListOrdersBWCNCHob9.Count == 0 || ListOrdersBWCNCHob10.Count == 0)
                        {
                            //2020-02-27 Start edit: đổi SharingBWOrdersToCNC() sang SharingBWOrdersToCNC2()
                            //chỉ chia đơn ko phải N hoặc N+1
                            SharingBWOrdersToCNC_NormalMode();
                            //2020-02-27 End edit

                        }
                        BalancedOrdersBW_NormalMode();
                    }
                        
                    //2020-02-24 End edit
                    //2019-12-19 End add

                    //2019-11-29 Start add: kiểm tra CNCHob9.json, CNCHob10.json để lấy ra TeethShape, TeethQuantity, Diameter_d cho việc sắp xếp các đơn của BWCNCHob9, BWCNCHob10
                    //cập nhật TeethShape, TeethQuantity, Diameter_d, DMat
                    Hob9BWCNC_TeethShape = orderCNCHob9.TeethShape;
                    Hob9BWCNC_TeethQty = orderCNCHob9.TeethQuantity;
                    Hob9BWCNC_Diameter = orderCNCHob9.Diameter_d;
                    Hob9BWCNC_DMat = orderCNCHob9.Material_text1;
                    Hob9BWCNC_GlobalCode = orderCNCHob9.Global_Code;

                    Hob10BWCNC_TeethShape = orderCNCHob10.TeethShape;
                    Hob10BWCNC_TeethQty = orderCNCHob10.TeethQuantity;
                    Hob10BWCNC_Diameter = orderCNCHob10.Diameter_d;
                    Hob10BWCNC_DMat = orderCNCHob10.Material_text1;
                    Hob10BWCNC_GlobalCode = orderCNCHob10.Global_Code;
                    //2019-11-29 End add
                }
                catch (Exception ex)
                {
                    MessageBox.Show("404 BW error: \n Không thể kết nối đồng bộ dữ liệu giữa BW và CNC. \n" + ex.Message,"Thông báo",MessageBoxButton.OK, MessageBoxImage.Error);
                }

              
            });

            timerCNCHob = new DispatcherTimer()
            {
                Interval = TimeSpan.FromSeconds(5)
            };

            //timerCNCHob.Start();
            timerCNCHob.Tick += new EventHandler((s, e) =>
            {
                try
                {
                    
                    OrderModel orderBWCNCHob9 = new OrderModel(MySQLProvider.Instance.ExecuteQuery("SELECT * FROM CNCHob9").Rows[0]);
                    OrderModel orderBWCNCHob10 = new OrderModel(MySQLProvider.Instance.ExecuteQuery("SELECT * FROM CNCHob10").Rows[0]);

                    //2020-02-08 Start add: lấy flag check BWCNCHob9, 10 kiểm tra nếu = 1 thì khóa ko cho in
                    var cnc9 = OrderBLL.Instance.GetLockStatusCNC9();
                    var cnc10  = OrderBLL.Instance.GetLockStatusCNC10();
                    var bw_cnc9_HasBWOrders = cnc9.Item2;
                    var bw_cnc10_HasBWOrders = cnc10.Item2;

                    //2020-03-23 Start add: thêm chức năng khóa in hàng F1 bên BW
                    var bw_cnc9_isLockedButton = cnc9.Item1;
                    var bw_cnc10_isLockedButton = cnc10.Item1;
                    //2020-03-23 End add
                    if (bw_cnc9_HasBWOrders == 1 || bw_cnc9_isLockedButton == 1)
                    {
                        //khóa nút in của CNCHob9 
                        //26-10-2020: tạm khóa
                        //CNCHob9_IsGettingData = false;
                        //26-10-2020
                    }
                    else
                    {
                        //nếu bw_cnc9_HasBWOrders = 0(tức là hết đơn BWCNCHob9 ko có đơn BW) và CNCHob9 có đơn => CNCHob9_IsGettingData = true;
                        if (ListOrdersCNCHob9 != null && ListOrdersCNCHob9.Count > 0)
                            CNCHob9_IsGettingData = true;
                    }

                    if(bw_cnc10_HasBWOrders == 1 || bw_cnc10_isLockedButton == 1)
                    {
                        //khóa nút in của CNCHob10 
                        //26-10-2020: tạm khóa
                        //CNCHob10_IsGettingData = false;
                        //26-10-2020 
                    }else
                    {
                        //nếu bw_cnc10_HasBWOrders = 0(tức là hết đơn BWCNCHob10 ko có đơn BW) và CNCHob10 có đơn => CNCHob10_IsGettingData = true;
                        if (ListOrdersCNCHob10 != null && ListOrdersCNCHob10.Count > 0)
                            CNCHob10_IsGettingData = true;
                    }

                    //2019-12-02 End change
                    if (orderBWCNCHob9 != null)
                    {
                        Hob9CNC_TeethShape = orderBWCNCHob9.TeethShape;
                        Hob9CNC_TeethQty = orderBWCNCHob9.TeethQuantity;
                        Hob9CNC_Diameter = orderBWCNCHob9.Diameter_d;
                        Hob9CNC_GlobalCode = orderBWCNCHob9.Global_Code;
                    }

                    if (orderBWCNCHob10 != null)
                    {
                        Hob10CNC_TeethShape = orderBWCNCHob10.TeethShape;
                        Hob10CNC_TeethQty = orderBWCNCHob10.TeethQuantity;
                        Hob10CNC_Diameter = orderBWCNCHob10.Diameter_d;
                        Hob10CNC_GlobalCode = orderBWCNCHob10.Global_Code;
                    }
                    //2019-11-29 End add

                    //2020-03-12 Start add: chia đều đơn N, N+1 cho 2 máy CNC
                    var ordersCNC9 = ListOrdersCNCHob9 != null ? ListOrdersCNCHob9 : (new ObservableCollection<OrderModel>());
                    var ordersCNC10 = ListOrdersCNCHob10 != null ? ListOrdersCNCHob10 : (new ObservableCollection<OrderModel>());
                    ApplySSDCheck_CNC(ordersCNC9.Union(ordersCNC10).ToList());

                    int status = OrderBLL.Instance.Get_cncIsSetupN1();
                    if (status == 1) //chia đơn N, N+1 đều ra các máy CNC
                    {
                        SharingOrdersCNC("SSDMode");
                    }
                  
                    //2020-03-12 End add

                    //2020-02-08 Start add: thêm function cân bằng order giữa 2 máy CNC
                    BalancedOrdersCNC2();
                    //2020-02-08 End add
                }
                catch (Exception ex)
                {
                    //2020-05-28 Start edit: tạm thời khóa
                    //MessageBox.Show("404 CNC error: \n Không thể kết nối đồng bộ dữ liệu giữa BW và CNC. \n" + ex.Message, "Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            });
            //2019-10-14

            //2019-11-14 Start add: add timer get data every 1 hour
            timerGetData = new DispatcherTimer()
            {
                //2019-12-11 Start change: đổi thời gian tự động lấy order mới mỗi 10 phút
                Interval = TimeSpan.FromMinutes(10)
                //2019-12-11 End add
            };

            timerGetData.Tick += (s, e) =>
            {
                if (IsGettingDataBW)
                {
                    if (btnGetDataBW != null)
                    {
                        typeof(System.Windows.Controls.Primitives.ButtonBase).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(btnGetDataBW, new object[0]);
                    }
                }

                if (IsGettingDataCNC)
                {
                    if (btnGetDataCNC != null)
                    {
                        typeof(System.Windows.Controls.Primitives.ButtonBase).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(btnGetDataCNC, new object[0]);
                    }
                }
            };
            //2019-11-14 End add
        }
        #endregion

        #region <<PROPERTIES>>
        //2019-10-16: KEY-BW-PROPERTIES
        //BW has 5 screens and only use 1 printer
        DispatcherTimer timerBWHob;
        DispatcherTimer timerBWNoHob;
        DispatcherTimer timerGetData;

        private ChildBWCNC_VM _childBWCNC_VM;
        public ChildBWCNC_VM Child_BWCNCVM
        {
            get { return _childBWCNC_VM; }
            set { _childBWCNC_VM = value;OnPropertyChanged(nameof(Child_BWCNCVM)); }
        }

        //2020-03-05 Start add: SSD mode check
        private string _ssdModeCheck_BW;
        public string SSDModeCheck_BW
        {
            get { return _ssdModeCheck_BW; }
            set
            {
                _ssdModeCheck_BW = value;
                OnPropertyChanged(nameof(SSDModeCheck_BW));
            }
        }

        private string _ssdModeCheck_CNC;
        public string SSDModeCheck_CNC
        {
            get { return _ssdModeCheck_CNC; }
            set
            {
                _ssdModeCheck_CNC = value;
                OnPropertyChanged(nameof(SSDModeCheck_CNC));
            }
        }
        //2020-03-05 End add
        Button btnGetDataBW;
        Button btnGetDataCNC;

        //2019-11-28 Start add: thêm điều kiện để gửi mail
        bool IsSendingMailBW = false;
        bool IsSendingMailCNC = false;
        //2019-11-28 End add

        WebBrowser _webBWHob11;
        WebBrowser _webBWHob12;
        WebBrowser _webBWCNCHob9;
        WebBrowser _webBWCNCHob10;

        WebBrowser _webCNCHob;

        bool isFirstNavigationBWNoHob = true;
        bool isFirstNavigationBWHob11 = true;
        bool isFirstNavigationBWHob12 = true;
        bool isFirstNavigationBWCNCHob9 = true;
        bool isFirstNavigationBWCNCHob10 = true;
        bool isFirstNavigationBWF1 = true;

        bool isSendMailSSDModeCheck_BW = false;
        bool isSendMailSSDModeCheck_CNC = false;
        //2019-11-25 Start add: thêm 2 thuộc tính kiểm tra Hob9BWCNC và Hob10BWCNC hết đơn
        bool BWCNCHob9_NeedOrders = false;
        bool BWCNCHob10_NeedOrders = false;
        //2019-11-25 End add

        //2019-11-30 Start add: thêm flag check BWCNC9, 10 có đơn từ BWHob11, 12 chuyển qua
        bool BWCNCHob9_HasBWOrders = false;
        bool BWCNCHob10_HasBWOrders = false;
        //2019-11-30 End add

        //2020-03-23 Start add: thêm flag để khóa/ mở khóa các nút in với đơn F1    
        private ICommand cmdLockPrint;
        public ICommand CommandLockPrint
        {
            get {
                if (cmdLockPrint == null)
                    cmdLockPrint = new RelayCommand<object>(CanLockPrintButton,LockPrintButton);

                return cmdLockPrint;
            }
            private set { cmdLockPrint = value; }
        }

      

        private string lockedButtonName;
        public string LockedButtonName
        {
            get { return lockedButtonName; }
            set { lockedButtonName = value; OnPropertyChanged(nameof(LockedButtonName)); }
        }
        private bool CanLockPrintButton(object obj)
        {
            if (string.IsNullOrEmpty(LockedButtonName))
                return false;
            else return true;
        }

        private void LockPrintButton(object obj)
        {
            string machine = obj.ToString();
            UnlockWindow wd = new UnlockWindow();
            Child_BWCNCVM = new ChildBWCNC_VM();
            Child_BWCNCVM.PropertyChanged += ChildBWCNC_VM_PropertyChanged;
            wd.DataContext = Child_BWCNCVM;            
            wd.Show();
           
        }

        private void ChildBWCNC_VM_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var child = sender as ChildBWCNC_VM;
            if(e.PropertyName == nameof(Child_BWCNCVM.PasswordUnlock))
            {
                var pass = OrderBLL.Instance.CheckUnlockSettings(Child_BWCNCVM.PasswordUnlock);
                if(pass == 1)
                {
                    if (LockedButtonName.Equals("11"))
                    {
                        if (BWHob11_IsGettingData)
                        {
                            BWHob11_IsGettingData = false;
                            BW_Hob11Status = "Máy 11 khóa để in order F1";
                        }
                        else
                        {
                            BWHob11_IsGettingData = true;
                            BW_Hob11Status = "Máy 11 mở khóa PRINT";
                        }
                    }
                    else if (LockedButtonName.Equals("12"))
                    {
                        if (BWHob12_IsGettingData)
                        {
                            BWHob12_IsGettingData = false;
                            BW_Hob12Status = "Máy 12 khóa PRINT để in order F1";
                        }
                        else
                        {
                            BWHob12_IsGettingData = true;
                            BW_Hob12Status = "Máy 12 mở khóa PRINT";
                        }
                    }
                    else if (LockedButtonName.Equals("9"))
                    {
                        if (BWCNCHob9_IsGettingData)
                        {
                            BWCNCHob9_IsGettingData = false;
                            OrderBLL.Instance.SetLockCNC9(1);
                            BWCNC_Hob9Status = "Máy 9 khóa PRINT để in order F1";
                        }
                        else
                        {
                            BWCNCHob9_IsGettingData = true;
                            OrderBLL.Instance.SetLockCNC9(0);
                            BWCNC_Hob9Status = "Máy 9 mở khóa PRINT";
                        }
                    }
                    else if (LockedButtonName.Equals("10"))
                    {
                        if (BWCNCHob10_IsGettingData)
                        {
                            BWCNCHob10_IsGettingData = false;
                            OrderBLL.Instance.SetLockCNC10(1);
                            BWCNC_Hob10Status = "Máy 10 khóa PRINT để in order F1";
                        }
                        else
                        {
                            BWCNCHob10_IsGettingData = true;
                            OrderBLL.Instance.SetLockCNC10(0);
                            BWCNC_Hob10Status = "Máy 10 mở khóa PRINT";
                        }
                    }
                }
               
            }
        }


        //2020-03-23 End add
        bool isFirstLoadBWNoHob = true;
        int pageNumberBWNoHob = 0;

        bool isFirstLoadBWF1 = true;
        int pageNumberBWF1 = 0;

        bool isFirstLoadBWHob11 = true;
        int pageNumberBWHob11 = 0;

        bool isFirstLoadBWHob12 = true;
        int pageNumberBWHob12 = 0;

        bool isFirstLoadBWCNCHob9 = true;
        int pageNumberBWCNCHob9 = 0;

        bool isFirstLoadBWCNCHob10 = true;
        int pageNumberBWCNCHob10 = 0;

        bool _BWHob11_IsOutOfSession = false;
        bool _BWHob12_IsOutOfSession = false;
        bool _BWCNCHob9_IsOutOfSession = false;
        bool _BWCNCHob10_IsOutOfSession = false;
        bool _BWNoHob_IsOutOfSession = false;
        bool _BWF1_IsOutOfSession = false;

        //2019-10-16

        //2019-10-16: KEY-CNC-PROPERTIES
        //CNC has 3 screens and only use 1 printer
        DispatcherTimer timerCNCHob;
        DispatcherTimer timerCNCNoHob;

        bool isFirstNavigationCNCNoHob = true;
        bool isFirstNavigationCNCHob = true;

        bool isFirstLoadCNCNoHob = true;
        int pageNumberCNCNoHob = 0;

        bool isFirstLoadCNCHob = true;
        int pageNumberCNCHob = 0;

        bool _CNCHob_IsOutOfSession = false;
        bool _CNCNoHob_IsOutOfSession = false;
        //2019-10-16

        //2019-10-16 Button GetData
        private bool _isLoadingBW;
        public bool IsLoadingBW
        {
            get { return _isLoadingBW; }
            set
            {
                _isLoadingBW = value;
                OnPropertyChanged(nameof(IsLoadingBW));
            }
        }

        private bool _IsGettingDataBW;
        public bool IsGettingDataBW
        {
            get { return _IsGettingDataBW; }
            set { _IsGettingDataBW = value; OnPropertyChanged(nameof(IsGettingDataBW)); }
        }

        private bool _isLoadingCNC;
        public bool IsLoadingCNC
        {
            get { return _isLoadingCNC; }
            set
            {
                _isLoadingCNC = value;
                OnPropertyChanged(nameof(IsLoadingCNC));
            }
        }

        private bool _IsGettingDataCNC;
        public bool IsGettingDataCNC
        {
            get { return _IsGettingDataCNC; }
            set { _IsGettingDataCNC = value; OnPropertyChanged(nameof(IsGettingDataCNC)); }
        }

        private ObservableCollection<OrderModel> _BWOrders;
        public ObservableCollection<OrderModel> BWOrders
        {
            get { return _BWOrders; }
            set { _BWOrders = value; OnPropertyChanged(nameof(BWOrders)); }
        }

        private ObservableCollection<OrderModel> _CNCOrders;
        public ObservableCollection<OrderModel> CNCOrders
        {
            get { return _CNCOrders; }
            set { _CNCOrders = value; OnPropertyChanged(nameof(CNCOrders)); }
        }
        //2019-10-16
        #region <<LINE-A-BW>>
        //2019-10-16 Line A BW
        private ObservableCollection<OrderModel> _listOrdersBWHob11;
        public ObservableCollection<OrderModel> ListOrdersBWHob11
        {
            get { return _listOrdersBWHob11; }
            set { _listOrdersBWHob11 = value; OnPropertyChanged(nameof(ListOrdersBWHob11)); }
        }

        private ObservableCollection<OrderModel> _listOrdersBWHob12;
        public ObservableCollection<OrderModel> ListOrdersBWHob12
        {
            get { return _listOrdersBWHob12; }
            set { _listOrdersBWHob12 = value; OnPropertyChanged(nameof(ListOrdersBWHob12)); }
        }

        private ObservableCollection<OrderModel> _listOrdersBWNoHob;
        public ObservableCollection<OrderModel> ListOrdersBWNoHob
        {
            get { return _listOrdersBWNoHob; }
            set { _listOrdersBWNoHob = value; OnPropertyChanged(nameof(ListOrdersBWNoHob)); }
        }

        private ObservableCollection<OrderModel> _listOrdersBWF1;
        public ObservableCollection<OrderModel> ListOrdersBWF1
        {
            get { return _listOrdersBWF1; }
            set { _listOrdersBWF1 = value; OnPropertyChanged(nameof(ListOrdersBWF1)); }
        }

        private OrderModel _selectedOrderBWHob11;
        public OrderModel SelectedOrderBWHob11
        {
            get { return _selectedOrderBWHob11; }
            set { _selectedOrderBWHob11 = value; OnPropertyChanged(nameof(SelectedOrderBWHob11)); }
        }

        private OrderModel _selectedOrderBWHob12;
        public OrderModel SelectedOrderBWHob12
        {
            get { return _selectedOrderBWHob12; }
            set { _selectedOrderBWHob12 = value; OnPropertyChanged(nameof(SelectedOrderBWHob12)); }
        }

        private OrderModel _selectedOrderBWNoHob;
        public OrderModel SelectedOrderBWNoHob
        {
            get { return _selectedOrderBWNoHob; }
            set { _selectedOrderBWNoHob = value; OnPropertyChanged(nameof(SelectedOrderBWNoHob)); }
        }

        private OrderModel _selectedOrderBWF1;
        public OrderModel SelectedOrderBWF1
        {
            get { return _selectedOrderBWF1; }
            set { _selectedOrderBWF1 = value; OnPropertyChanged(nameof(SelectedOrderBWF1)); }
        }

        private bool _BWHob11_IsPrinting;
        public bool BWHob11_IsPrinting
        {
            get { return _BWHob11_IsPrinting; }
            set { _BWHob11_IsPrinting = value; OnPropertyChanged(nameof(BWHob11_IsPrinting)); }
        }

        private bool _BWHob12_IsPrinting;
        public bool BWHob12_IsPrinting
        {
            get { return _BWHob12_IsPrinting; }
            set { _BWHob12_IsPrinting = value; OnPropertyChanged(nameof(BWHob12_IsPrinting)); }
        }

        private bool _BWNoHob_IsPrinting;
        public bool BWNoHob_IsPrinting
        {
            get { return _BWNoHob_IsPrinting; }
            set { _BWNoHob_IsPrinting = value; OnPropertyChanged(nameof(BWNoHob_IsPrinting)); }
        }

        private bool _BWF1_IsPrinting;
        public bool BWF1_IsPrinting
        {
            get { return _BWF1_IsPrinting; }
            set { _BWF1_IsPrinting = value; OnPropertyChanged(nameof(BWF1_IsPrinting)); }
        }

        private bool _BWHob11_IsGettingData;
        public bool BWHob11_IsGettingData
        {
            get { return _BWHob11_IsGettingData; }
            set { _BWHob11_IsGettingData = value; OnPropertyChanged(nameof(BWHob11_IsGettingData)); }
        }

        private bool _BWHob12_IsGettingData;
        public bool BWHob12_IsGettingData
        {
            get { return _BWHob12_IsGettingData; }
            set { _BWHob12_IsGettingData = value; OnPropertyChanged(nameof(BWHob12_IsGettingData)); }
        }

        private bool _BWNoHob_IsGettingData;
        public bool BWNoHob_IsGettingData
        {
            get { return _BWNoHob_IsGettingData; }
            set { _BWNoHob_IsGettingData = value; OnPropertyChanged(nameof(BWNoHob_IsGettingData)); }
        }

        private bool _BWF1_IsGettingData;
        public bool BWF1_IsGettingData
        {
            get { return _BWF1_IsGettingData; }
            set { _BWF1_IsGettingData = value; OnPropertyChanged(nameof(BWF1_IsGettingData)); }
        }

        public string Hob11BW_TeethShape
        {
            get { return Properties.Settings.Default._hob11BW_TeethShape; }
            set
            {
                Properties.Settings.Default._hob11BW_TeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob11BW_TeethShape));
            }
        }

        public int Hob11BW_TeethQty
        {
            get { return Properties.Settings.Default._hob11BW_TeethQuantity; }
            set
            {
                Properties.Settings.Default._hob11BW_TeethQuantity = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob11BW_TeethQty));
            }
        }

        //2019-11-29 Start add: thêm Diameter
        public double Hob11BW_Diameter
        {
            get { return Properties.Settings.Default._hob11BW_Diameter; }
            set
            {
                Properties.Settings.Default._hob11BW_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob11BW_Diameter));
            }
        }

        public string Hob11BW_GlobalCode
        {
            get { return Properties.Settings.Default._hob11BW_GlobalCode; }
            set
            {
                Properties.Settings.Default._hob11BW_GlobalCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob11BW_GlobalCode));
            }
        }

        public double Hob12BW_Diameter
        {
            get { return Properties.Settings.Default._hob12BW_Diameter; }
            set
            {
                Properties.Settings.Default._hob12BW_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob12BW_Diameter));
            }
        }

        public string Hob12BW_GlobalCode
        {
            get { return Properties.Settings.Default._hob12BW_GlobalCode; }
            set
            {
                Properties.Settings.Default._hob12BW_GlobalCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob12BW_GlobalCode));
            }
        }

        public double NoHobBW_Diameter
        {
            get { return Properties.Settings.Default._noHobBW_Diameter; }
            set
            {
                Properties.Settings.Default._noHobBW_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(NoHobBW_Diameter));
            }
        }
        //2019-11-29 End add

        public string Hob12BW_TeethShape
        {
            get { return Properties.Settings.Default._hob12BW_TeethShape; }
            set
            {
                Properties.Settings.Default._hob12BW_TeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob12BW_TeethShape));
            }
        }

        public int Hob12BW_TeethQty
        {
            get { return Properties.Settings.Default._hob12BW_TeethQuantity; }
            set
            {
                Properties.Settings.Default._hob12BW_TeethQuantity = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob12BW_TeethQty));
            }
        }

        public string NoHobBW_TeethShape
        {
            get { return Properties.Settings.Default._noHobBWTeethShape; }
            set
            {
                Properties.Settings.Default._noHobBWTeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(NoHobBW_TeethShape));
            }
        }

        public int NoHobBW_TeethQty
        {
            get { return Properties.Settings.Default._noHobCNCTeethQty; }
            set
            {
                Properties.Settings.Default._noHobCNCTeethQty = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(NoHobBW_TeethQty));
            }
        }

        private string _BW_Hob11Status;
        public string BW_Hob11Status
        {
            get { return _BW_Hob11Status; }
            set { _BW_Hob11Status = value; OnPropertyChanged(nameof(BW_Hob11Status)); }
        }

        private string _BW_Hob12Status;
        public string BW_Hob12Status
        {
            get { return _BW_Hob12Status; }
            set { _BW_Hob12Status = value; OnPropertyChanged(nameof(BW_Hob12Status)); }
        }

        private string _BW_NoHobStatus;
        public string BW_NoHobStatus
        {
            get { return _BW_NoHobStatus; }
            set { _BW_NoHobStatus = value; OnPropertyChanged(nameof(BW_NoHobStatus)); }
        }

        private string _BW_F1Status;
        public string BW_F1Status
        {
            get { return _BW_F1Status; }
            set { _BW_F1Status = value; OnPropertyChanged(nameof(BW_F1Status)); }
        }

        public double F1BW_Diameter
        {
            get { return Properties.Settings.Default._noHobCNCTeethQty; }
            set
            {
                Properties.Settings.Default._f1BW_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(F1BW_Diameter));
            }
        }
        //2019-10-16

        //2019-11-27 Start add: total PO and total orders
        private int _totalPO_BWHob11;
        public int TotalPO_BWHob11 { get { return _totalPO_BWHob11; } set { _totalPO_BWHob11 = value; OnPropertyChanged(nameof(TotalPO_BWHob11)); } }

        private int _totalOrders_BWHob11;
        public int TotalOrders_BWHob11 { get { return _totalOrders_BWHob11; } set { _totalOrders_BWHob11 = value; OnPropertyChanged(nameof(TotalOrders_BWHob11)); } }

        private int _totalPO_BWHob12;
        public int TotalPO_BWHob12 { get { return _totalPO_BWHob12; } set { _totalPO_BWHob12 = value; OnPropertyChanged(nameof(TotalPO_BWHob12)); } }

        private int _totalOrders_BWHob12;
        public int TotalOrders_BWHob12 { get { return _totalOrders_BWHob12; } set { _totalOrders_BWHob12 = value; OnPropertyChanged(nameof(TotalOrders_BWHob12)); } }

        private int _totalPO_BWCNCHob9;
        public int TotalPO_BWCNCHob9 { get { return _totalPO_BWCNCHob9; } set { _totalPO_BWCNCHob9 = value; OnPropertyChanged(nameof(TotalPO_BWCNCHob9)); } }

        private int _totalOrders_BWCNCHob9;
        public int TotalOrders_BWCNCHob9 { get { return _totalOrders_BWCNCHob9; } set { _totalOrders_BWCNCHob9 = value; OnPropertyChanged(nameof(TotalOrders_BWCNCHob9)); } }

        private int _totalPO_BWCNCHob10;
        public int TotalPO_BWCNCHob10 { get { return _totalPO_BWCNCHob10; } set { _totalPO_BWCNCHob10 = value; OnPropertyChanged(nameof(TotalPO_BWCNCHob10)); } }

        private int _totalOrders_BWCNCHob10;
        public int TotalOrders_BWCNCHob10 { get { return _totalOrders_BWCNCHob10; } set { _totalOrders_BWCNCHob10 = value; OnPropertyChanged(nameof(TotalOrders_BWCNCHob10)); } }

        private int _totalPO_BWNoHob;
        public int TotalPO_BWNoHob { get { return _totalPO_BWNoHob; } set { _totalPO_BWNoHob = value; OnPropertyChanged(nameof(TotalPO_BWNoHob)); } }

        private int _totalOrders_BWNoHob;
        public int TotalOrders_BWNoHob { get { return _totalOrders_BWNoHob; } set { _totalOrders_BWNoHob = value; OnPropertyChanged(nameof(TotalOrders_BWNoHob)); } }

        private int _totalPO_BWF1;
        public int TotalPO_BWF1 { get { return _totalPO_BWF1; } set { _totalPO_BWF1 = value; OnPropertyChanged(nameof(TotalPO_BWF1)); } }

        private int _totalOrders_BWF1;
        public int TotalOrders_BWF1 { get { return _totalOrders_BWF1; } set { _totalOrders_BWF1 = value; OnPropertyChanged(nameof(TotalOrders_BWF1)); } }

        //2019-12-05 Start add: trạng thái unlock, lock setting BW      
        private bool _BW_IsLockSettings;
        public bool BW_IsUnLockSettings { get { return _BW_IsLockSettings; } set { _BW_IsLockSettings = value; OnPropertyChanged(nameof(BW_IsUnLockSettings)); } }

        private string _accountStatusBW;
        public string AccountStatusBW
        {
            get { return _accountStatusBW; }
            set { _accountStatusBW = value; OnPropertyChanged(nameof(AccountStatusBW)); }
        }

        private Visibility _BW_BlockIsVisible;
        public Visibility BW_BlockIsVisible
        {
            get { return _BW_BlockIsVisible; }
            set { _BW_BlockIsVisible = value; OnPropertyChanged(nameof(BW_BlockIsVisible)); }
        }

        private string _BW_AccountPassword;
        public string BW_AccountPassword
        {
            get { return _BW_AccountPassword; }
            set { _BW_AccountPassword = value; OnPropertyChanged(nameof(BW_AccountPassword)); }
        }
        //2019-12-05 End add

        //2020-03-09 Start add:
        private double _remainPT_BW;
        public double BW_RemainPT
        {
            get { return _remainPT_BW; }
            set { _remainPT_BW = value; OnPropertyChanged(nameof(BW_RemainPT)); }
        }

        private int _BW_Total_N1_Orders;
        public int BW_Total_N1_Orders
        {
            get { return _BW_Total_N1_Orders; }
            set { _BW_Total_N1_Orders = value; OnPropertyChanged(nameof(BW_Total_N1_Orders)); }
        }

        //2020-03-09 End add

        //2019-11-27 End add
        #endregion <<LINE-A-BW>>

        #region <<LINE-B-BW&CNC>>
        //2019-10-16 Line B BW&CNC
        private ObservableCollection<OrderModel> _listOrdersBWCNCHob9;
        public ObservableCollection<OrderModel> ListOrdersBWCNCHob9
        {
            get { return _listOrdersBWCNCHob9; }
            set { _listOrdersBWCNCHob9 = value; OnPropertyChanged(nameof(ListOrdersBWCNCHob9)); }
        }

        private ObservableCollection<OrderModel> _listOrdersBWCNCHob10;
        public ObservableCollection<OrderModel> ListOrdersBWCNCHob10
        {
            get { return _listOrdersBWCNCHob10; }
            set { _listOrdersBWCNCHob10 = value; OnPropertyChanged(nameof(ListOrdersBWCNCHob10)); }
        }

        private OrderModel _selectedOrderBWCNCHob9;
        public OrderModel SelectedOrderBWCNCHob9
        {
            get { return _selectedOrderBWCNCHob9; }
            set { _selectedOrderBWCNCHob9 = value; OnPropertyChanged(nameof(SelectedOrderBWCNCHob9)); }
        }

        private OrderModel _selectedOrderBWCNCHob10;
        public OrderModel SelectedOrderBWCNCHob10
        {
            get { return _selectedOrderBWCNCHob10; }
            set { _selectedOrderBWCNCHob10 = value; OnPropertyChanged(nameof(SelectedOrderBWCNCHob10)); }
        }

        private bool _BWCNCHob9_IsPrinting;
        public bool BWCNCHob9_IsPrinting
        {
            get { return _BWCNCHob9_IsPrinting; }
            set { _BWCNCHob9_IsPrinting = value; OnPropertyChanged(nameof(BWCNCHob9_IsPrinting)); }
        }

        private bool _BWCNCHob10_IsPrinting;
        public bool BWCNCHob10_IsPrinting
        {
            get { return _BWCNCHob10_IsPrinting; }
            set { _BWCNCHob10_IsPrinting = value; OnPropertyChanged(nameof(BWCNCHob10_IsPrinting)); }
        }

        private bool _BWCNCHob9_IsGettingData;
        public bool BWCNCHob9_IsGettingData
        {
            get { return _BWCNCHob9_IsGettingData; }
            set { _BWCNCHob9_IsGettingData = value; OnPropertyChanged(nameof(BWCNCHob9_IsGettingData)); }
        }

        private bool _BWCNCHob10_IsGettingData;
        public bool BWCNCHob10_IsGettingData
        {
            get { return _BWCNCHob10_IsGettingData; }
            set { _BWCNCHob10_IsGettingData = value; OnPropertyChanged(nameof(BWCNCHob10_IsGettingData)); }
        }

        //2019-11-29 Start add: thêm đường kính cho BWCNCHob9, BWCNCHob10
        public double Hob9BWCNC_Diameter
        {
            get { return Properties.Settings.Default._hob9BWCNC_Diameter; }
            set
            {
                Properties.Settings.Default._hob9BWCNC_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9BWCNC_Diameter));
            }
        }

        public double Hob10BWCNC_Diameter
        {
            get { return Properties.Settings.Default._hob10BWCNC_Diameter; }
            set
            {
                Properties.Settings.Default._hob10BWCNC_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10BWCNC_Diameter));
            }
        }
        //2019-11-29 End add        
        public string Hob9BWCNC_TeethShape
        {
            get { return Properties.Settings.Default._hob9BWCNC_TeethShape; }
            set
            {
                Properties.Settings.Default._hob9BWCNC_TeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9BWCNC_TeethShape));
            }
        }

        public string Hob9BWCNC_DMat
        {
            get { return Properties.Settings.Default._hob9BWCNC_DMat; }
            set
            {
                Properties.Settings.Default._hob9BWCNC_DMat = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9BWCNC_DMat));
            }
        }

        public int Hob9BWCNC_TeethQty
        {
            get { return Properties.Settings.Default._hob9BWCNC_TeethQuantity; }
            set
            {
                Properties.Settings.Default._hob9BWCNC_TeethQuantity = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9BWCNC_TeethQty));
            }
        }

        public string Hob9BWCNC_GlobalCode
        {
            get { return Properties.Settings.Default._hob9BWCNC_GlobalCode; }
            set
            {
                Properties.Settings.Default._hob9BWCNC_GlobalCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9BWCNC_GlobalCode));
            }
        }

        public string Hob10BWCNC_TeethShape
        {
            get { return Properties.Settings.Default._hob10BWCNC_TeethShape; }
            set
            {
                Properties.Settings.Default._hob10BWCNC_TeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10BWCNC_TeethShape));
            }
        }
        public string Hob10BWCNC_DMat
        {
            get { return Properties.Settings.Default._hob10BWCNC_DMat; }
            set
            {
                Properties.Settings.Default._hob10BWCNC_DMat = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10BWCNC_DMat));
            }
        }

        public int Hob10BWCNC_TeethQty
        {
            get { return Properties.Settings.Default._hob10BWCNC_TeethQuantity; }
            set
            {
                Properties.Settings.Default._hob10BWCNC_TeethQuantity = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10BWCNC_TeethQty));
            }
        }

        public string Hob10BWCNC_GlobalCode
        {
            get { return Properties.Settings.Default._hob10BWCNC_GlobalCode; }
            set
            {
                Properties.Settings.Default._hob10BWCNC_GlobalCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10BWCNC_GlobalCode));
            }
        }

        private string _BWCNC_Hob9Status;
        public string BWCNC_Hob9Status
        {
            get { return _BWCNC_Hob9Status; }
            set { _BWCNC_Hob9Status = value; OnPropertyChanged(nameof(BWCNC_Hob9Status)); }
        }

        private string _BWCNC_Hob10Status;
        public string BWCNC_Hob10Status
        {
            get { return _BWCNC_Hob10Status; }
            set { _BWCNC_Hob10Status = value; OnPropertyChanged(nameof(BWCNC_Hob10Status)); }
        }
        //2019-10-16
        //2019-12-03 Start add:
        private bool _Hob9BWCNC_TeethShape_IsEnabled;
        public bool Hob9BWCNC_TeethShape_IsEnabled
        {
            get { return _Hob9BWCNC_TeethShape_IsEnabled; }
            set { _Hob9BWCNC_TeethShape_IsEnabled = value; OnPropertyChanged(nameof(Hob9BWCNC_TeethShape_IsEnabled)); }
        }

        private bool _Hob9BWCNC_TeethQty_IsEnabled;
        public bool Hob9BWCNC_TeethQty_IsEnabled
        {
            get { return _Hob9BWCNC_TeethQty_IsEnabled; }
            set { _Hob9BWCNC_TeethQty_IsEnabled = value; OnPropertyChanged(nameof(Hob9BWCNC_TeethQty_IsEnabled)); }
        }

        private bool _Hob9BWCNC_Diameter_IsEnabled;
        public bool Hob9BWCNC_Diameter_IsEnabled
        {
            get { return _Hob9BWCNC_Diameter_IsEnabled; }
            set { _Hob9BWCNC_Diameter_IsEnabled = value; OnPropertyChanged(nameof(Hob9BWCNC_Diameter_IsEnabled)); }
        }

        private bool _Hob9BWCNC_GlobalCode_IsEnabled;
        public bool Hob9BWCNC_GlobalCode_IsEnabled
        {
            get { return _Hob9BWCNC_GlobalCode_IsEnabled; }
            set { _Hob9BWCNC_GlobalCode_IsEnabled = value; OnPropertyChanged(nameof(Hob9BWCNC_GlobalCode_IsEnabled)); }
        }

        private bool _Hob10BWCNC_TeethShape_IsEnabled;
        public bool Hob10BWCNC_TeethShape_IsEnabled
        {
            get { return _Hob10BWCNC_TeethShape_IsEnabled; }
            set { _Hob10BWCNC_TeethShape_IsEnabled = value; OnPropertyChanged(nameof(Hob10BWCNC_TeethShape_IsEnabled)); }
        }

        private bool _Hob10BWCNC_TeethQty_IsEnabled;
        public bool Hob10BWCNC_TeethQty_IsEnabled
        {
            get { return _Hob10BWCNC_TeethQty_IsEnabled; }
            set { _Hob10BWCNC_TeethQty_IsEnabled = value; OnPropertyChanged(nameof(Hob10BWCNC_TeethQty_IsEnabled)); }
        }

        private bool _Hob10BWCNC_Diameter_IsEnabled;
        public bool Hob10BWCNC_Diameter_IsEnabled
        {
            get { return _Hob10BWCNC_Diameter_IsEnabled; }
            set { _Hob10BWCNC_Diameter_IsEnabled = value; OnPropertyChanged(nameof(Hob10BWCNC_Diameter_IsEnabled)); }
        }

        private bool _Hob10BWCNC_GlobalCode_IsEnabled;
        public bool Hob10BWCNC_GlobalCode_IsEnabled
        {
            get { return _Hob10BWCNC_GlobalCode_IsEnabled; }
            set { _Hob10BWCNC_GlobalCode_IsEnabled = value; OnPropertyChanged(nameof(Hob10BWCNC_GlobalCode_IsEnabled)); }
        }
        //2019-12-03 End add
        #endregion <<LINE-B-BW&CNC>>

        #region <<LINE-B-CNC>>
        //2019-10-16 Line B CNC
        private ObservableCollection<OrderModel> _listOrdersCNCHob9;
        public ObservableCollection<OrderModel> ListOrdersCNCHob9
        {
            get { return _listOrdersCNCHob9; }
            set { _listOrdersCNCHob9 = value; OnPropertyChanged(nameof(ListOrdersCNCHob9)); }
        }

        private ObservableCollection<OrderModel> _listOrdersCNCHob10;
        public ObservableCollection<OrderModel> ListOrdersCNCHob10
        {
            get { return _listOrdersCNCHob10; }
            set { _listOrdersCNCHob10 = value; OnPropertyChanged(nameof(ListOrdersCNCHob10)); }
        }

        private ObservableCollection<OrderModel> _listOrdersCNCNoHob;
        public ObservableCollection<OrderModel> ListOrdersCNCNoHob
        {
            get { return _listOrdersCNCNoHob; }
            set { _listOrdersCNCNoHob = value; OnPropertyChanged(nameof(ListOrdersCNCNoHob)); }
        }

        private OrderModel _selectedOrderCNCHob9;
        public OrderModel SelectedOrderCNCHob9
        {
            get { return _selectedOrderCNCHob9; }
            set { _selectedOrderCNCHob9 = value; OnPropertyChanged(nameof(SelectedOrderCNCHob9)); }
        }

        private OrderModel _selectedOrderCNCHob10;
        public OrderModel SelectedOrderCNCHob10
        {
            get { return _selectedOrderCNCHob10; }
            set { _selectedOrderCNCHob10 = value; OnPropertyChanged(nameof(SelectedOrderCNCHob10)); }
        }

        private OrderModel _selectedOrderCNCNoHob;
        public OrderModel SelectedOrderCNCNoHob
        {
            get { return _selectedOrderCNCNoHob; }
            set { _selectedOrderCNCNoHob = value; OnPropertyChanged(nameof(SelectedOrderCNCNoHob)); }
        }

        private bool _CNCHob9_IsPrinting;
        public bool CNCHob9_IsPrinting
        {
            get { return _CNCHob9_IsPrinting; }
            set { _CNCHob9_IsPrinting = value; OnPropertyChanged(nameof(CNCHob9_IsPrinting)); }
        }

        private bool _CNCHob10_IsPrinting;
        public bool CNCHob10_IsPrinting
        {
            get { return _CNCHob10_IsPrinting; }
            set { _CNCHob10_IsPrinting = value; OnPropertyChanged(nameof(CNCHob10_IsPrinting)); }
        }

        private bool _CNCNoHob_IsPrinting;
        public bool CNCNoHob_IsPrinting
        {
            get { return _CNCNoHob_IsPrinting; }
            set { _CNCNoHob_IsPrinting = value; OnPropertyChanged(nameof(CNCNoHob_IsPrinting)); }
        }

        private bool _CNCHob9_IsGettingData;
        public bool CNCHob9_IsGettingData
        {
            get { return _CNCHob9_IsGettingData; }
            set { _CNCHob9_IsGettingData = value; OnPropertyChanged(nameof(CNCHob9_IsGettingData)); }
        }

        private bool _CNCHob10_IsGettingData;
        public bool CNCHob10_IsGettingData
        {
            get { return _CNCHob10_IsGettingData; }
            set { _CNCHob10_IsGettingData = value; OnPropertyChanged(nameof(CNCHob10_IsGettingData)); }
        }

        private bool _CNCNoHob_IsGettingData;
        public bool CNCNoHob_IsGettingData
        {
            get { return _CNCNoHob_IsGettingData; }
            set { _CNCNoHob_IsGettingData = value; OnPropertyChanged(nameof(CNCNoHob_IsGettingData)); }
        }

        //2019-11-29 Start add: thêm đường kính CNCHob9, CNCHob10
        public double Hob9CNC_Diameter
        {
            get { return Properties.Settings.Default._hob9CNC_Diameter; }
            set
            {
                Properties.Settings.Default._hob9CNC_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9CNC_Diameter));
            }
        }

        public double Hob10CNC_Diameter
        {
            get { return Properties.Settings.Default._hob10CNC_Diameter; }
            set
            {
                Properties.Settings.Default._hob10CNC_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10CNC_Diameter));
            }
        }

        public double NoHobCNC_Diameter
        {
            get { return Properties.Settings.Default._noHobCNC_Diameter; }
            set
            {
                Properties.Settings.Default._noHobCNC_Diameter = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(NoHobCNC_Diameter));
            }
        }
        //2019-11-29 End add

        public string Hob9CNC_TeethShape
        {
            get { return Properties.Settings.Default._hob9CNC_TeethShape; }
            set
            {
                Properties.Settings.Default._hob9CNC_TeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9CNC_TeethShape));
            }
        }

        public int Hob9CNC_TeethQty
        {
            get { return Properties.Settings.Default._hob9CNC_TeethQuantity; }
            set
            {
                Properties.Settings.Default._hob9CNC_TeethQuantity = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9CNC_TeethQty));
            }
        }

        public string Hob9CNC_DMat
        {
            get { return Properties.Settings.Default._hob9CNC_DMat; }
            set
            {
                Properties.Settings.Default._hob9CNC_DMat = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9CNC_DMat));
            }
        }

        public string Hob9CNC_GlobalCode
        {
            get { return Properties.Settings.Default._hob9CNC_GlobalCode; }
            set
            {
                Properties.Settings.Default._hob9CNC_GlobalCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob9CNC_GlobalCode));
            }
        }

        public string Hob10CNC_TeethShape
        {
            get { return Properties.Settings.Default._hob10CNC_TeethShape; }
            set
            {
                Properties.Settings.Default._hob10CNC_TeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10CNC_TeethShape));
            }
        }

        public int Hob10CNC_TeethQty
        {
            get { return Properties.Settings.Default._hob10CNC_TeethQuantity; }
            set
            {
                Properties.Settings.Default._hob10CNC_TeethQuantity = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10CNC_TeethQty));
            }
        }

        public string Hob10CNC_DMat
        {
            get { return Properties.Settings.Default._hob10CNC_DMat; }
            set
            {
                Properties.Settings.Default._hob10CNC_DMat = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10CNC_DMat));
            }
        }

        public string Hob10CNC_GlobalCode
        {
            get { return Properties.Settings.Default._hob10CNC_GlobalCode; }
            set
            {
                Properties.Settings.Default._hob10CNC_GlobalCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(Hob10CNC_GlobalCode));
            }
        }

        public string NoHobCNC_TeethShape
        {
            get { return Properties.Settings.Default._noHobCNCTeethShape; }
            set
            {
                Properties.Settings.Default._noHobCNCTeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(NoHobCNC_TeethShape));
            }
        }

        public int NoHobCNC_TeethQty
        {
            get { return Properties.Settings.Default._noHobCNCTeethQty; }
            set
            {
                Properties.Settings.Default._noHobCNCTeethQty = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(NoHobCNC_TeethQty));
            }
        }

        private string _CNC_Hob9Status;
        public string CNC_Hob9Status
        {
            get { return _CNC_Hob9Status; }
            set { _CNC_Hob9Status = value; OnPropertyChanged(nameof(CNC_Hob9Status)); }
        }

        private string _CNC_Hob10Status;
        public string CNC_Hob10Status
        {
            get { return _CNC_Hob10Status; }
            set { _CNC_Hob10Status = value; OnPropertyChanged(nameof(CNC_Hob10Status)); }
        }

        private string _CNC_NoHobStatus;
        public string CNC_NoHobStatus
        {
            get { return _CNC_NoHobStatus; }
            set { _CNC_NoHobStatus = value; OnPropertyChanged(nameof(CNC_NoHobStatus)); }
        }
        //2019-10-16

        //2019-11-27 Start add: total PO and total orders
        private int _totalPO_CNCHob9;
        public int TotalPO_CNCHob9 { get { return _totalPO_CNCHob9; } set { _totalPO_CNCHob9 = value; OnPropertyChanged(nameof(TotalPO_CNCHob9)); } }

        private int _totalOrders_CNCHob9;
        public int TotalOrders_CNCHob9 { get { return _totalOrders_CNCHob9; } set { _totalOrders_CNCHob9 = value; OnPropertyChanged(nameof(TotalOrders_CNCHob9)); } }

        private int _totalPO_CNCHob10;
        public int TotalPO_CNCHob10 { get { return _totalPO_CNCHob10; } set { _totalPO_CNCHob10 = value; OnPropertyChanged(nameof(TotalPO_CNCHob10)); } }

        private int _totalOrders_CNCHob10;
        public int TotalOrders_CNCHob10 { get { return _totalOrders_CNCHob10; } set { _totalOrders_CNCHob10 = value; OnPropertyChanged(nameof(TotalOrders_CNCHob10)); } }

        private int _totalPO_CNCNoHob;
        public int TotalPO_CNCNoHob { get { return _totalPO_CNCNoHob; } set { _totalPO_CNCNoHob = value; OnPropertyChanged(nameof(TotalPO_CNCNoHob)); } }

        private int _totalOrders_CNCNoHob;
        public int TotalOrders_CNCNoHob { get { return _totalOrders_CNCNoHob; } set { _totalOrders_CNCNoHob = value; OnPropertyChanged(nameof(TotalOrders_CNCNoHob)); } }

        //2019-12-05 Start add: trạng thái unlock, lock setting CNC      
        private bool _CNC_IsUnLockSettings;
        public bool CNC_IsUnLockSettings { get { return _CNC_IsUnLockSettings; } set { _CNC_IsUnLockSettings = value; OnPropertyChanged(nameof(CNC_IsUnLockSettings)); } }

        private string _accountStatusCNC;
        public string AccountStatusCNC
        {
            get { return _accountStatusCNC; }
            set { _accountStatusCNC = value; OnPropertyChanged(nameof(AccountStatusCNC)); }
        }

        private Visibility _CNC_BlockIsVisible;
        public Visibility CNC_BlockIsVisible
        {
            get { return _CNC_BlockIsVisible; }
            set { _CNC_BlockIsVisible = value; OnPropertyChanged(nameof(CNC_BlockIsVisible)); }
        }

        private string _CNC_AccountPassword;
        public string CNC_AccountPassword
        {
            get { return _CNC_AccountPassword; }
            set { _CNC_AccountPassword = value; OnPropertyChanged(nameof(CNC_AccountPassword)); }
        }
        //2019-12-05 End add
        //2019-11-27 End add

        //2019-12-03 Start add:
        private bool _Hob9CNC_TeethShape_IsEnabled;
        public bool Hob9CNC_TeethShape_IsEnabled
        {
            get { return _Hob9CNC_TeethShape_IsEnabled; }
            set { _Hob9CNC_TeethShape_IsEnabled = value; OnPropertyChanged(nameof(Hob9CNC_TeethShape_IsEnabled)); }
        }

        private bool _Hob9CNC_TeethQty_IsEnabled;
        public bool Hob9CNC_TeethQty_IsEnabled
        {
            get { return _Hob9CNC_TeethQty_IsEnabled; }
            set { _Hob9CNC_TeethQty_IsEnabled = value; OnPropertyChanged(nameof(Hob9CNC_TeethQty_IsEnabled)); }
        }

        private bool _Hob9CNC_Diameter_IsEnabled;
        public bool Hob9CNC_Diameter_IsEnabled
        {
            get { return _Hob9CNC_Diameter_IsEnabled; }
            set { _Hob9CNC_Diameter_IsEnabled = value; OnPropertyChanged(nameof(Hob9CNC_Diameter_IsEnabled)); }
        }

        private bool _Hob9CNC_Global_IsEnabled;
        public bool Hob9CNC_Global_IsEnabled
        {
            get { return _Hob9CNC_Global_IsEnabled; }
            set { _Hob9CNC_Global_IsEnabled = value; OnPropertyChanged(nameof(Hob9CNC_Global_IsEnabled)); }
        }

        private bool _Hob10CNC_TeethShape_IsEnabled;
        public bool Hob10CNC_TeethShape_IsEnabled
        {
            get { return _Hob10CNC_TeethShape_IsEnabled; }
            set { _Hob10CNC_TeethShape_IsEnabled = value; OnPropertyChanged(nameof(Hob10CNC_TeethShape_IsEnabled)); }
        }

        private bool _Hob10CNC_TeethQty_IsEnabled;
        public bool Hob10CNC_TeethQty_IsEnabled
        {
            get { return _Hob10CNC_TeethQty_IsEnabled; }
            set { _Hob10CNC_TeethQty_IsEnabled = value; OnPropertyChanged(nameof(Hob10CNC_TeethQty_IsEnabled)); }
        }

        private bool _Hob10CNC_Diameter_IsEnabled;
        public bool Hob10CNC_Diameter_IsEnabled
        {
            get { return _Hob10CNC_Diameter_IsEnabled; }
            set { _Hob10CNC_Diameter_IsEnabled = value; OnPropertyChanged(nameof(Hob10CNC_Diameter_IsEnabled)); }
        }

        private bool _Hob10CNC_Global_IsEnabled;
        public bool Hob10CNC_Global_IsEnabled
        {
            get { return _Hob10CNC_Global_IsEnabled; }
            set { _Hob10CNC_Global_IsEnabled = value; OnPropertyChanged(nameof(Hob10CNC_Global_IsEnabled)); }
        }
        //2019-12-03 End add

        //2020-03-12 Start add: thêm thông số check SSD, N+1 Qty cho CNC
        private double _CNC_RemainPT;
        public double CNC_RemainPT
        {
            get { return _CNC_RemainPT; }
            set { _CNC_RemainPT = value; OnPropertyChanged(nameof(CNC_RemainPT)); }
        }

        private int _CNC_Total_N1_Orders;
        public int CNC_Total_N1_Orders
        {
            get { return _CNC_Total_N1_Orders; }
            set { _CNC_Total_N1_Orders = value; OnPropertyChanged(nameof(CNC_Total_N1_Orders)); }
        }
        //2020-03-12 End add
        #endregion <<LINE-B-CNC>>

        #region <<MAIL-PROPERTIES>>
        //2019-10-16 Mail properties
        public string SPCMailSender
        {
            get
            {
                return Properties.Settings.Default._spcMailSender;
            }
            set
            {
                Properties.Settings.Default._spcMailSender = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(SPCMailSender));
            }
        }


        public string SPCMailReceiver
        {
            get { return Properties.Settings.Default._spcMailReceiver; }
            set
            {
                Properties.Settings.Default._spcMailReceiver = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(SPCMailReceiver));
            }
        }


        public string SPCMailCode
        {
            get { return Properties.Settings.Default._spcMailCode; }
            set
            {
                Properties.Settings.Default._spcMailCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(SPCMailCode));
            }
        }
        //2019-10-16
        #endregion <<MAIL-PROPERTIES>>

        #endregion <<PROPERTIES>>

        #region <<COMMANDS>>



        #region <<BW-COMMANDS>>
        private ICommand _commandGetOrdersBW;
        public ICommand CommandGetOrdersBW
        {
            get
            {
                if (_commandGetOrdersBW == null)
                    _commandGetOrdersBW = new RelayCommand<object>(CanGetOrdersBW, GetOrdersBW);

                return _commandGetOrdersBW;
            }
            private set { _commandGetOrdersBW = value; }
        }

        private ICommand _commandGetOrdersCNC;
        public ICommand CommandGetOrdersCNC
        {
            get
            {
                if (_commandGetOrdersCNC == null)
                    _commandGetOrdersCNC = new RelayCommand<object>(CanGetOrdersCNC, GetOrdersCNC);

                return _commandGetOrdersCNC;
            }
            private set { _commandGetOrdersCNC = value; }
        }


        private ICommand _commandBWPrintHob11;
        public ICommand CommandBWPrintHob11
        {
            get
            {
                if (_commandBWPrintHob11 == null)
                    _commandBWPrintHob11 = new RelayCommand<object>(CanPrintBWHob11, PrintBWHob11);
                return _commandBWPrintHob11;
            }
            private set { _commandBWPrintHob11 = value; }
        }

        private ICommand _commandBWPrintHob12;
        public ICommand CommandBWPrintHob12
        {
            get
            {
                if (_commandBWPrintHob12 == null)
                    _commandBWPrintHob12 = new RelayCommand<object>(CanPrintBWHob12, PrintBWHob12);
                return _commandBWPrintHob12;
            }
            private set { _commandBWPrintHob12 = value; }
        }

        private ICommand _commandBWPrintNoHob;
        public ICommand CommandBWPrintNoHob
        {
            get
            {
                if (_commandBWPrintNoHob == null)
                    _commandBWPrintNoHob = new RelayCommand<object>(CanPrintBWNoHob, PrintBWNoHob);
                return _commandBWPrintNoHob;
            }
            private set { _commandBWPrintNoHob = value; }
        }

        private ICommand _commandBWPrintF1;
        public ICommand CommandBWPrintF1
        {
            get
            {
                if (_commandBWPrintF1 == null)
                    _commandBWPrintF1 = new RelayCommand<object>(CanPrintBWF1, PrintBWF1);
                return _commandBWPrintF1;
            }
            private set { _commandBWPrintF1 = value; }
        }

        private ICommand _commandBWCNCPrintHob9;
        public ICommand CommandBWCNCPrintHob9
        {
            get
            {
                if (_commandBWCNCPrintHob9 == null)
                    _commandBWCNCPrintHob9 = new RelayCommand<object>(CanPrintBWCNCHob9, PrintBWCNCHob9);
                return _commandBWCNCPrintHob9;
            }
            private set { _commandBWCNCPrintHob9 = value; }
        }



        private ICommand _commandBWCNCPrintHob10;
        public ICommand CommandBWCNCPrintHob10
        {
            get
            {
                if (_commandBWCNCPrintHob10 == null)
                    _commandBWCNCPrintHob10 = new RelayCommand<object>(CanPrintBWCNCHob10, PrintBWCNCHob10);
                return _commandBWCNCPrintHob10;
            }
            private set { _commandBWCNCPrintHob10 = value; }
        }

        //2019-12-03 Start add: thêm toggle on/off sửa TeethShape, TeethQty, Diameter_d của BWCNCHob9, BWCNCHob10
        private ICommand _cmdToggleBWCNCHob9;
        public ICommand CommandToggleBWCNCHob9
        {
            get
            {
                if (_cmdToggleBWCNCHob9 == null)
                    _cmdToggleBWCNCHob9 = new RelayCommand<object>(ToggleOnOff_BWCNCHob9);
                return _cmdToggleBWCNCHob9;
            }
            private set { _cmdToggleBWCNCHob9 = value; }
        }

        private ICommand _cmdToggleBWCNCHob10;
        public ICommand CommandToggleBWCNCHob10
        {
            get
            {
                if (_cmdToggleBWCNCHob10 == null)
                    _cmdToggleBWCNCHob10 = new RelayCommand<object>(ToggleOnOff_BWCNCHob10);
                return _cmdToggleBWCNCHob10;
            }
            private set { _cmdToggleBWCNCHob10 = value; }
        }

        //2019-12-03 End add

        //2019-12-05 Start add: command Lock, Unlock settings
        private ICommand _cmdBWUnLockSettings;
        public ICommand CommandBWUnLockSettings
        {
            get
            {
                if (_cmdBWUnLockSettings == null)
                    _cmdBWUnLockSettings = new RelayCommand<object>(BWUnLockSettings);
                return _cmdBWUnLockSettings;
            }
            private set { _cmdBWUnLockSettings = value; }
        }

        private ICommand _cmdBWOpenSettings;
        public ICommand CommandBWOpenSettings
        {
            get
            {
                if (_cmdBWOpenSettings == null)
                    _cmdBWOpenSettings = new RelayCommand<object>(BWOpenSettings);
                return _cmdBWOpenSettings;
            }
            private set { _cmdBWOpenSettings = value; }
        }

        private void BWOpenSettings(object obj)
        {
            //kiểm tra password, có => check MYSQL => open settings, bỏ block che datagrid, ko => vẫn khóa
          
               
                int result = 0;
                try
                {
                    result = OrderBLL.Instance.CheckUnlockSettings(BW_AccountPassword);
                    if (result > 0)
                    {
                        AccountStatusBW = "Mở khóa cài đặt thành công!";
                        BW_IsUnLockSettings = true;
                        BW_BlockIsVisible = Visibility.Hidden;
                        BW_AccountPassword = "";
                    }
                    else
                    {
                        AccountStatusBW = "Mở khóa cài đặt thất bại, vui lòng kiểm tra mật khẩu và thử lại!";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Something went wrong!\n"+ex.Message,"Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                }
          
        }

        private void BWUnLockSettings(object obj)
        {
            if (Hob9BWCNC_TeethShape_IsEnabled || Hob10BWCNC_TeethShape_IsEnabled)
            {
                MessageBox.Show("Vui lòng lưu thông số sắp xếp orders ở máy CNC9, CNC10 trước khi khóa cài đặt","Thông báo",MessageBoxButton.OK, MessageBoxImage.Warning);
            }else
            {
                AccountStatusBW = "Cài đặt đã khóa!";
                BW_IsUnLockSettings = false;
                BW_BlockIsVisible = Visibility.Visible;
            }
        }

        //2019-12-05 End add
        #endregion <<BW-COMMANDS>>

        #region <<CNC-COMMANDS>>
        private ICommand _commandCNCPrintHob9;
        public ICommand CommandCNCPrintHob9
        {
            get
            {
                if (_commandCNCPrintHob9 == null)
                    _commandCNCPrintHob9 = new RelayCommand<object>(CanPrintCNCHob9, PrintCNCHob9);
                return _commandCNCPrintHob9;
            }
            private set { _commandCNCPrintHob9 = value; }
        }


        private ICommand _commandCNCPrintHob10;
        public ICommand CommandCNCPrintHob10
        {
            get
            {
                if (_commandCNCPrintHob10 == null)
                    _commandCNCPrintHob10 = new RelayCommand<object>(CanPrintCNCHob10, PrintCNCHob10);
                return _commandCNCPrintHob10;
            }
            private set { _commandCNCPrintHob10 = value; }
        }


        private ICommand _commandCNCPrintNoHob;
        public ICommand CommandCNCPrintNoHob
        {
            get
            {
                if (_commandCNCPrintNoHob == null)
                    _commandCNCPrintNoHob = new RelayCommand<object>(CanPrintCNCNoHob, PrintCNCNoHob);
                return _commandCNCPrintNoHob;
            }
            private set { _commandCNCPrintNoHob = value; }

        }

        //2019-12-03 Start add: thêm toggle on/off edit TeethShape, TeethQty, Diameter_d của CNCHob9, CNCHob10
        private ICommand _cmdToggleCNCHob9;
        public ICommand CommandToggleCNCHob9
        {
            get
            {
                if (_cmdToggleCNCHob9 == null)
                    _cmdToggleCNCHob9 = new RelayCommand<object>(ToggleOnOff_CNCHob9);
                return _cmdToggleCNCHob9;
            }
            private set { _cmdToggleCNCHob9 = value; }
        }

        private ICommand _cmdToggleCNCHob10;
        public ICommand CommandToggleCNCHob10
        {
            get
            {
                if (_cmdToggleCNCHob10 == null)
                    _cmdToggleCNCHob10 = new RelayCommand<object>(ToggleOnOff_CNCHob10);
                return _cmdToggleCNCHob10;
            }
            private set { _cmdToggleCNCHob10 = value; }
        }


        //2019-12-03 End add

        //2019-12-05 Start add: command Lock, Unlock settings
        private ICommand _cmdCNCUnLockSettings;
        public ICommand CommandCNCUnLockSettings
        {
            get
            {
                if (_cmdCNCUnLockSettings == null)
                    _cmdCNCUnLockSettings = new RelayCommand<object>(CNCUnLockSettings);
                return _cmdCNCUnLockSettings;
            }
            private set { _cmdCNCUnLockSettings = value; }
        }

        private ICommand _cmdCNCOpenSettings;
        public ICommand CommandCNCOpenSettings
        {
            get
            {
                if (_cmdCNCOpenSettings == null)
                    _cmdCNCOpenSettings = new RelayCommand<object>(CNCOpenSettings);
                return _cmdCNCOpenSettings;
            }
            private set { _cmdCNCOpenSettings = value; }
        }

        private void CNCOpenSettings(object obj)
        {
            //kiểm tra password, có => check MYSQL => open settings, bỏ block che datagrid, ko => vẫn khóa
           
                
                int result = 0;
                try
                {
                    result = OrderBLL.Instance.CheckUnlockSettings(CNC_AccountPassword);
                    if (result > 0)
                    {
                        AccountStatusBW = "Mở khóa cài đặt thành công!";
                        CNC_IsUnLockSettings = true;
                        CNC_BlockIsVisible = Visibility.Hidden;
                        CNC_AccountPassword = "";
                    }
                    else
                    {
                        AccountStatusCNC = "Mở khóa cài đặt thất bại, vui lòng kiểm tra mật khẩu và thử lại!";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Something went wrong!\n" + ex.Message, "Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                }

           
        }

        private void CNCUnLockSettings(object obj)
        {
            if (Hob9CNC_TeethShape_IsEnabled || Hob10CNC_TeethShape_IsEnabled)
            {
                MessageBox.Show("Vui lòng lưu thông số sắp xếp orders ở máy CNC9, CNC10 trước khi khóa cài đặt", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                AccountStatusCNC = "Cài đặt đã khóa!";
                CNC_IsUnLockSettings = false;
                CNC_BlockIsVisible = Visibility.Visible;
            }
        }

        //2019-12-05 End add
        #endregion <<CNC-COMMANDS>>

        #endregion <<COMMANDS>>

        #region <<FUNCTIONS>>
        //KEY-BW-1
        

        WebBrowser wbTest;
        //KEY-BW-2: BWHob11_LoadComplete        
        private void wbBWHob11_LoadCompleted(object sender, NavigationEventArgs e)
        {
            isFirstNavigationBWHob11 = false;
            WebBrowser wb = sender as WebBrowser;
            mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

            if (isFirstLoadBWHob11)
            {

                //get user login info in App.config
                string user = ConfigurationManager.AppSettings["user"];
                string pass = ConfigurationManager.AppSettings["pass"];

                //login
                mshtml.IHTMLElement userInput = document.getElementById("txtUserId");
                mshtml.IHTMLElement passInput = document.getElementById("txtPassword");
                mshtml.IHTMLElement btnLogin = document.getElementById("cmdLogin");
                userInput.setAttribute("value", user);
                passInput.setAttribute("value", pass);
                btnLogin.click();
                isFirstLoadBWHob11 = false;
                //điều hướng trang về 0 để vào trang in 
                pageNumberBWHob11 = 0;
            }
            else
            {

                if (pageNumberBWHob11 == 0)
                {
                    //nếu trang tải xong
                    if (wb.IsLoaded)
                    {
                        //check session time out
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadBWHob11 = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {
                            //still on the session
                            var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                            var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                            script.text = @" 
                                   window.onload = function(){
                                        //navigate 'Manufacturing Instructions Print' menu
                                        window.location = '/manufa/PrintViewInstruction.aspx';
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);
                            HideScriptErrors(wb, true);
                            pageNumberBWHob11 = 1;
                        }
                    }

                }
                else if (pageNumberBWHob11 == 1)
                {
                    //check session time out
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if BWHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _BWHob11_IsOutOfSession = true;
                        //2019-10-11 End add

                        //navigate to login page
                        isFirstLoadBWHob11 = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();

                    }
                    else
                    {
                        //2019-10-11 Start add
                        if (_BWHob11_IsOutOfSession)
                        {
                            //_BWHob11_IsOutOfSession = false;//reset to default value
                            if (BWHob11_IsPrinting) //nếu BWHob11 đang in
                            {
                                SelectedOrderBWHob11 = ListOrdersBWHob11[0];
                            }
                            else if (BWHob12_IsPrinting) //nếu BWHob12 đang in
                            {
                                SelectedOrderBWHob12 = ListOrdersBWHob12[0];
                            }
                            else if (BWCNCHob9_IsPrinting) //nếu BWCNCHob9 đang in
                            {
                                SelectedOrderBWCNCHob9 = ListOrdersBWCNCHob9[0];
                            }
                            else if (BWCNCHob10_IsPrinting)
                            {
                                SelectedOrderBWCNCHob10 = ListOrdersBWCNCHob10[0];
                            }

                            //onclick btnClear to navigate to page 3
                            pageNumberBWHob11 = 2;
                            var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                            btnClear.click();
                        }
                        //2019-10-11 End add
                        else
                        {
                            var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                            var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                            script.text = @" 
                                   window.onload = function(){                                                                              
                                        //set selected status option for the order that not printed yet
                                        var printMode = document.getElementById('ContentPlaceHolder1_rdoSelectOne');
                                        printMode.checked = true;        
                                          
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);
                            pageNumberBWHob11 = 2;

                        }

                    }

                }
                else if (pageNumberBWHob11 == 2)
                {
                    //2019-11-14 Start add: cập nhật lại trạng thái session và trạng thái nút in sau khi re-login, thông báo in lại
                    if (_BWHob11_IsOutOfSession)
                    {
                        _BWHob11_IsOutOfSession = false;
                        IsGettingDataBW = true;
                        if (BWHob11_IsPrinting)
                        {
                            BWHob11_IsGettingData = true;
                            BWHob11_IsPrinting = false;
                            BW_Hob11Status = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                        }
                        else if (BWHob12_IsPrinting)
                        {
                            BWHob12_IsGettingData = true;
                            BWHob12_IsPrinting = false;
                            BW_Hob12Status = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                        }
                        else if (BWCNCHob9_IsPrinting)
                        {
                            BWCNCHob9_IsGettingData = true;
                            BWCNCHob9_IsPrinting = false;
                            BWCNC_Hob9Status = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                        }
                        else if (BWCNCHob10_IsPrinting)
                        {
                            BWCNCHob10_IsGettingData = true;
                            BWCNCHob10_IsPrinting = false;
                            BWCNC_Hob10Status = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                        }

                    }
                    
                }
                else if (pageNumberBWHob11 == 3)
                {
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if BWHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _BWHob11_IsOutOfSession = true;
                        //2019-10-11 End add

                        //thông báo hết session
                        if (BWHob11_IsPrinting)
                            BW_Hob11Status = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                        else if (BWHob12_IsPrinting)
                            BW_Hob12Status = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                        else if (BWCNCHob9_IsPrinting)
                            BWCNC_Hob9Status = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                        else if (BWCNCHob10_IsPrinting)
                            BWCNC_Hob10Status = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                        //navigate to login page
                        isFirstLoadBWHob11 = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();

                    }
                    else
                    {

                        var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                        var script = (mshtml.IHTMLScriptElement)document.createElement("script");

                        script.text = @"window.onload = function(){                
                                                                                    
                                        var orderToPrint = document.getElementById('ContentPlaceHolder1_grdPrintList_chkOutput_0');
                                        var btnPrint = document.getElementById('ContentPlaceHolder1_cmdPrint');
                                            
                                        if(orderToPrint != null)
                                        {   
                                            //check vào PO cần in
                                            orderToPrint.checked = 'checked';   
                                            //chọn máy in, 15 = pulley máy 2
                                            var printer = document.getElementById('ContentPlaceHolder1_drpPrinter');
                                            //2019-12-05 Manufa thay đổi thứ tự máy in
                                            //printer.selectedIndex = 15 
                                            printer.value = 121; //máy pulley 2 value = 121
                                            //2019-12-05 
                                            //in 
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },500);
                                            
                                            
                                        }
                                   }";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        //2019-11-06 Start lock
                        pageNumberBWHob11 = 4; //chuyển trang sau khi in, lấy status
                                               //2019-11-06 End lock

                    }

                }
                else if (pageNumberBWHob11 == 4)
                {
                    //2019-11-06 Start add: kiểm tra PO được in hay ko
                    if (wb.IsLoaded)
                    {
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadBWHob11 = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {
                            //lấy status sau khi in
                            string status = document.getElementById("ContentPlaceHolder1_AlertArea").innerHTML;
                            //nếu in thành công
                            if (status.ToLower().Contains("completed"))
                            {
                                if (BWHob11_IsPrinting) //nếu BWHob11 đang in
                                {
                                    Hob11BW_TeethShape = SelectedOrderBWHob11.TeethShape;
                                    Hob11BW_TeethQty = SelectedOrderBWHob11.TeethQuantity;
                                    Hob11BW_Diameter = SelectedOrderBWHob11.Diameter_d;
                                    Hob11BW_GlobalCode = SelectedOrderBWHob11.Global_Code;

                                    //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                    var poToDelete = ListOrdersBWHob11.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWHob11.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        //cập nhật 3 thông số trên lên server để đồng bộ
                                        OrderBLL.Instance.Update_BW11_TeethInfo(poToDelete);

                                        ListOrdersBWHob11.Remove(poToDelete);
                                        //2019-12-02 Start lock: ko cần lưu thông tin PO của BWHob11 vì ko cần đồng bộ thông tin
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWHob11.json");                                        
                                        //2019-12-02 End lock
                                        BW_Hob11Status = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                        Task.Delay(1000).Wait();
                                        //2019-10-15 Start add: enable timer again
                                        timerBWHob.Start();
                                        //2019-10-15 End add
                                    }
                              
                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable BWHobLoading while the printer is running
                                    IsGettingDataBW = true;
                                    BWHob11_IsGettingData = true;
                                    BWHob11_IsPrinting = false;
                                    //2019-10-07 End add
                                    //2019-11-05 End lock

                                 
                                }
                                else if (BWHob12_IsPrinting) //nếu BWHob12 đang in
                                {
                                    Hob12BW_TeethShape = SelectedOrderBWHob12.TeethShape;
                                    Hob12BW_TeethQty = SelectedOrderBWHob12.TeethQuantity;
                                    Hob12BW_Diameter = SelectedOrderBWHob12.Diameter_d;
                                    Hob12BW_GlobalCode = SelectedOrderBWHob12.Global_Code;
                                    //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                    var poToDelete = ListOrdersBWHob12.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWHob12.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        //cập nhật 3 thông số trên lên server để đồng bộ
                                        OrderBLL.Instance.Update_BW12_TeethInfo(poToDelete);

                                        ListOrdersBWHob12.Remove(poToDelete);
                                        //2019-12-02 Start lock: ko cần lưu PO của BWHob12 vì ko cần đồng bộ thông tin
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWHob12.json");
                                        //2019-12-02 End lock
                                        BW_Hob12Status = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                        Task.Delay(1000).Wait();
                                        //2019-10-15 Start add: enable timer again
                                        timerBWHob.Start();
                                        //2019-10-15 End add
                                    }
                                   

                                    //2019-11-05 Start lock
                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable BWHobLoading while the printer is running
                                    IsGettingDataBW = true;
                                    BWHob12_IsGettingData = true;
                                    BWHob12_IsPrinting = false;
                                    //2019-10-07 End add
                                    //2019-11-05 End lock
                                    
                                }
                                else if (BWCNCHob9_IsPrinting) //nếu BWCNCHob9 đang in
                                {
                                    Hob9BWCNC_TeethShape = SelectedOrderBWCNCHob9.TeethShape;
                                    Hob9BWCNC_TeethQty = SelectedOrderBWCNCHob9.TeethQuantity;
                                    Hob9BWCNC_Diameter = SelectedOrderBWCNCHob9.Diameter_d;
                                    Hob9BWCNC_GlobalCode = SelectedOrderBWCNCHob9.Global_Code;
                                     //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                     var poToDelete = ListOrdersBWCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWCNCHob9.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersBWCNCHob9.Remove(poToDelete);
                                        //2019-12-02 Start change: đổi cách lưu đồng bộ PO cho CNCHob9 -> lưu vào MYSQL
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWCNCHob9.json");
                                        string query = string.Format("UPDATE CNCHob9 SET SoPO='{0}', TeethShape='{1}', TeethQty={2}, Diameter_d={3}, GlobalCode = '{4}'",poToDelete.Manufa_Instruction_No, poToDelete.TeethShape, poToDelete.TeethQuantity, poToDelete.Diameter_d, poToDelete.Global_Code);
                                        try
                                        {
                                            MySQLProvider.Instance.ExecuteNonQuery(query);
                                            BWCNC_Hob9Status = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                            Task.Delay(1000).Wait();
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("In PO thất bại, vui lòng thử lại!","Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                        }
                                        //2019-12-02 End change
                                        
                                        //2019-10-15 Start add: enable timer again
                                        timerBWHob.Start();
                                        //2019-10-15 End add
                                    }
                                    //2019-11-04 Start lock
                                    //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                    //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    //btnClear.click();
                                    //2019-10-04 End add
                                    //2019-11-04 End lock

                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable BWHobLoading while the printer is running
                                    IsGettingDataBW = true;
                                    BWCNCHob9_IsGettingData = true;
                                    BWCNCHob9_IsPrinting = false;
                                    //2019-10-07 End add

                                    //2019-11-04 Start add
                                    //pageNumberBWHob11 = 5;//in tiếp
                                    //2019-11-04 End add
                                }
                                else if (BWCNCHob10_IsPrinting)
                                {
                                    Hob10BWCNC_TeethShape = SelectedOrderBWCNCHob10.TeethShape;
                                    Hob10BWCNC_TeethQty = SelectedOrderBWCNCHob10.TeethQuantity;
                                    Hob10BWCNC_Diameter = SelectedOrderBWCNCHob10.Diameter_d;
                                    Hob10BWCNC_GlobalCode = SelectedOrderBWCNCHob10.Global_Code;
                                    //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                    var poToDelete = ListOrdersBWCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWCNCHob10.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersBWCNCHob10.Remove(poToDelete);
                                        //2019-12-02 Start change: đổi cách lưu đồng bộ PO cho CNCHob10 -> lưu vào MYSQL
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWCNCHob10.json");
                                        string query = string.Format("UPDATE CNCHob10 SET SoPO='{0}', TeethShape='{1}', TeethQty={2}, Diameter_d={3}, GlobalCode={4}", poToDelete.Manufa_Instruction_No, poToDelete.TeethShape, poToDelete.TeethQuantity, poToDelete.Diameter_d, poToDelete.Global_Code);
                                        try
                                        {
                                            MySQLProvider.Instance.ExecuteNonQuery(query);
                                            BWCNC_Hob10Status = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                            Task.Delay(1000).Wait();
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("In PO thất bại, vui lòng thử lại!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                        }
                                        //2019-12-02 End change
                                                                                
                                        //2019-10-15 Start add: enable timer again
                                        timerBWHob.Start();
                                        //2019-10-15 End add
                                    }

                                    //2019-11-04 Start lock
                                    //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                    //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    //btnClear.click();
                                    //2019-10-04 End add
                                    //2019-11-04 End lock

                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable BWHobLoading while the printer is running
                                    IsGettingDataBW = true;
                                    BWCNCHob10_IsGettingData = true;
                                    BWCNCHob10_IsPrinting = false;
                                    //2019-10-07 End add

                                    //2019-11-04 Start add
                                    //pageNumberBWHob11 = 5;//in tiếp
                                    //2019-11-04 End add
                                }

                            }
                            else if (status.ToLower().Contains("found"))
                            {
                                //Ko tìm thấy PO
                                //PO chua duoc chon
                                if (BWHob11_IsPrinting) //nếu BWHob11 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersBWHob11.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWHob11.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersBWHob11.Remove(poToDelete);
                                        //2019-10-11 Start add: print message: Order has been printed
                                        BW_Hob11Status = "PO: " + poToDelete + " has been printed!";
                                    }
                                    //2019-10-07 End add
                                    BWHob11_IsGettingData = true;
                                    BWHob11_IsPrinting = false;
                                }
                                else if (BWHob12_IsPrinting) //nếu BWHob12 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersBWHob12.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWHob12.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersBWHob12.Remove(poToDelete);
                                        //2019-10-11 Start add: print message: Order has been printed
                                        BW_Hob12Status = "PO: " + poToDelete + " has been printed!";
                                    }
                                    //2019-10-07 End add
                                    BWHob12_IsGettingData = true;
                                    BWHob12_IsPrinting = false;
                                }
                                else if (BWCNCHob9_IsPrinting) //nếu BWCNCHob9 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersBWCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWCNCHob9.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersBWCNCHob9.Remove(poToDelete);
                                        //2019-10-11 Start add: print message: Order has been printed
                                        BWCNC_Hob9Status = "PO: " + poToDelete + " has been printed!";
                                    }
                                    //2019-10-07 End add
                                    BWCNCHob9_IsGettingData = true;
                                    BWCNCHob9_IsPrinting = false;
                                }
                                else if (BWCNCHob10_IsPrinting)
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersBWCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWCNCHob10.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersBWCNCHob10.Remove(poToDelete);
                                        //2019-10-11 Start add: print message: Order has been printed
                                        BWCNC_Hob9Status = "PO: " + poToDelete + " has been printed!";
                                    }
                                    //2019-10-07 End add
                                    BWCNCHob10_IsGettingData = true;
                                    BWCNCHob10_IsPrinting = false;
                                }

                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add


                                //Enable btnGetData
                                IsGettingDataBW = true;

                                //2019-10-07 End add
                                //chuyển hướng trang 
                                pageNumberBWHob11 = -1;
                            }
                            else
                            {
                                //máy in hết giấy, offline, etc
                                //quay về trang in
                                //2019-10-07 Start add
                                //Enable btnGetData, disable BWHobLoading if getting error
                                IsGettingDataBW = true;
                                if (BWHob11_IsPrinting) //nếu BWHob11 đang in
                                {
                                    //2019-10-11 Start add: print message: Order has been printed
                                    BW_Hob11Status = "Printer offline or paper ran out, please check and try again!";
                                    //2019-10-07 End add
                                    BWHob11_IsGettingData = true;
                                    BWHob11_IsPrinting = false;
                                }
                                else if (BWHob12_IsPrinting) //nếu BWHob12 đang in
                                {
                                    //2019-10-11 Start add: print message: Order has been printed
                                    BW_Hob12Status = "Printer offline or paper ran out, please check and try again!";
                                    //2019-10-07 End add
                                    BWHob12_IsGettingData = true;
                                    BWHob12_IsPrinting = false;
                                }
                                else if (BWCNCHob9_IsPrinting) //nếu BWCNCHob9 đang in
                                {
                                    //2019-10-11 Start add: print message: Order has been printed
                                    BWCNC_Hob9Status = "Printer offline or paper ran out, please check and try again!";
                                    //2019-10-07 End add
                                    BWCNCHob9_IsGettingData = true;
                                    BWCNCHob9_IsPrinting = false;
                                }
                                else if (BWCNCHob10_IsPrinting)
                                {
                                    //2019-10-11 Start add: print message: Order has been printed
                                    BWCNC_Hob10Status = "Printer offline or paper ran out, please check and try again!";
                                    //2019-10-07 End add
                                    BWCNCHob10_IsGettingData = true;
                                    BWCNCHob10_IsPrinting = false;
                                }


                                //2019-10-07 End add

                                //chuyển hướng trang 
                                pageNumberBWHob11 = -1;
                            }
                            //2019-11-27 Start add: calculate total PO and total Orders after print BW
                            Calc_PO_and_Orders_BW();
                            //2019-11-27 End add
                        }
                    }
                    //2019-11-06 End add
                   
                }
                else if (pageNumberBWHob11 == 5)
                {
                    //2019-11-05 Start add: Khi manufa lấy data về thì đưa PO sang Manufa website tìm kiếm sau đó check vào PO được tìm để chuẩn bị in
                    var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                    var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                    script.text = @" 
                                   window.onload = function(){                
                                                                                    
                                        var orderToPrint = document.getElementById('ContentPlaceHolder1_grdPrintList_chkOutput_0');
                                        var btnPrint = document.getElementById('ContentPlaceHolder1_cmdPrint');                                            
                                        if(orderToPrint != null)
                                        {                                               
                                            //check vào PO cần in
                                            orderToPrint.checked = 'checked';   
                                            //chọn máy in, 4 = pulley
                                            var printer = document.getElementById('ContentPlaceHolder1_drpPrinter');
                                            printer.selectedIndex = 4
                                                                                     
                                        }
                                   }
                              ";
                    head.appendChild((mshtml.IHTMLDOMNode)script);
                    //2019-11-05 End add

                    //2019-11-07 Start add: 

                    //trả về trạng thái default cho btnPRINT sẵn sàng cho việc in PO tiếp theo
                    IsGettingDataBW = true;
                    if (BWHob11_IsPrinting)
                    {
                        BWHob11_IsGettingData = true;
                        BWHob11_IsPrinting = false;
                    }
                    else if (BWHob12_IsPrinting)
                    {
                        BWHob12_IsGettingData = true;
                        BWHob12_IsPrinting = false;
                    }
                    else if (BWCNCHob9_IsPrinting)
                    {
                        BWCNCHob9_IsGettingData = true;
                        BWCNCHob9_IsPrinting = false;
                    }
                    else if (BWCNCHob10_IsPrinting)
                    {
                        BWCNCHob10_IsGettingData = true;
                        BWCNCHob10_IsPrinting = false;
                    }
                    else if (BWNoHob_IsPrinting)
                    {
                        BWNoHob_IsGettingData = true;
                        BWNoHob_IsPrinting = false;
                    }
                    else if (BWF1_IsPrinting)
                    {
                        BWF1_IsGettingData = true;
                        BWF1_IsPrinting = false;
                    }
                    //2019-11-07 End add
                }
                else if (pageNumberBWHob11 == 100)
                {
                    if (wb.IsLoaded)
                    {
                        //2019-11-06 Start add
                        var table = document.getElementById("ContentPlaceHolder1_grdPrintList");
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(table.outerHTML);
                        var rows = doc.GetElementbyId("ContentPlaceHolder1_grdPrintList").SelectNodes("//*[@id='ContentPlaceHolder1_grdPrintList']/tbody/tr");
                        if (rows.Count < 2)
                        {
                            //if (BWHob11_IsPrinting)
                            //{
                            //    if(ListOrdersBWHob11 != null && ListOrdersBWHob11.Count > 0)//vẫn còn item
                            //    {
                            //        SelectedOrderBWHob11 = ListOrdersBWHob11.FirstOrDefault();
                            //        document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob11.Manufa_Instruction_No);
                            //        document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            //        //điều hướng trang
                            //        pageNumberBWHob = 3;
                            //    }
                            //    else
                            //    {
                            //        //hết item trong List=> ấn clear, trả lại trạng thái default cho btnPRINT

                            //        IsGettingDataBW = true;
                            //        BWHob11_IsGettingData = true;
                            //        BWHob11_IsPrinting = false;
                            //    }

                            //}
                            ////BWHob12,...

                            //điều hướng trang 
                            pageNumberBWHob11 = 3;
                        }
                        else
                        {
                            document.getElementById("ContentPlaceHolder1_cmdPrint").click();
                            //điều hướng trang
                            pageNumberBWHob11 = 3;
                        }
                    }

                }

            }

        }


       
        #region <<COMMAND-FUNCTIONS>>
        //2019-10-31 Start add
        int _mnfBWpage = 0;
        bool _mnfBW_getData = false;
        //2019-10-31 End add
        private async void GetOrdersBW(object parameters)
        {
            if (parameters != null)
            {
                await Task.Run(async () =>
                {

                    IsLoadingBW = true;
                    IsGettingDataBW = false;
                    timerBWHob.Stop();

                    ListOrdersBWHob11 = new ObservableCollection<OrderModel>();
                    ListOrdersBWHob12 = new ObservableCollection<OrderModel>();
                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>();
                    ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>();
                    ListOrdersBWNoHob = new ObservableCollection<OrderModel>();
                    ListOrdersBWF1 = new ObservableCollection<OrderModel>();

                    //2019-12-11 Start add: khi get Orders từ Manufa, disable tất cả các nút in
                    BWHob11_IsGettingData = false;
                    BWHob12_IsGettingData = false;
                    BWNoHob_IsGettingData = false;
                    BWCNCHob9_IsGettingData = false;
                    BWCNCHob10_IsGettingData = false;
                    //BWF1_IsGettingData = false;
                    //2019-12-11 End add

                    var values = parameters as object[];
              
                    //2019-11-07 Start add: sử dụng UserControl với Webbrowser 
                    UserControl ucBWHob11 = values[0] as UserControl;                  
                    //2019-11-07 End add
                  
                    Awesomium.Windows.Controls.WebControl wbBWNoHob = values[1] as Awesomium.Windows.Controls.WebControl;

                    //2019-11-14 Start add: gán button GetData để auto click lấy data mỗi giờ
                    timerGetData.Stop();
                    btnGetDataBW = values[3] as Button;
                    //2019-11-14 End add
                    await Task.Run(() =>
                    {
                        if (isFirstNavigationBWNoHob)
                        {
                            
                            //2019-11-11 Start add: dùng Awesomium thay cho Webbrowser
                            wbBWNoHob.Dispatcher.InvokeAsync(() =>
                            {
                                wbBWNoHob.Source = new Uri("http://10.4.24.111:8080/manufa/Login.aspx");
                                wbBWNoHob.LoadingFrameComplete += new Awesomium.Core.FrameEventHandler(wbBWNoHob_LoadedFrameComplete);
                            });

                        }

                        if (isFirstNavigationBWHob11)
                        {
                            ucBWHob11.Dispatcher.InvokeAsync(() =>
                            {
                                WebBrowser wbBWHob11 = ucBWHob11.FindName("wbBWHob11") as WebBrowser;
                                //navigate BWHob11 page
                                wbBWHob11.Dispatcher.InvokeAsync(() =>
                                {
                                    wbBWHob11.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                                    wbBWHob11.LoadCompleted += new LoadCompletedEventHandler(wbBWHob11_LoadCompleted);

                                    //2019-11-05 Start add: gán wbBWHob11 cho _webBWHob để điền PO mặc định khi dữ liệu Manufa được chuẩn bị xong
                                    _webBWHob11 = wbBWHob11;
                                    //2019-11-05
                                });
                            });

                        }

                       

                        //2019-10-31 Start add: đổi cách lấy dữ liệu manufa

                        //lấy manufa data cho BW
                        _mnfBW_getData = true;

                        WebBrowser wbManufaDT = values[2] as WebBrowser;
                        wbManufaDT.Dispatcher.InvokeAsync(() =>
                        {
                            if (wbManufaDT != null)
                                wbManufaDT.LoadCompleted -= ManufaDT_LoadCompleted;

                            wbManufaDT.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");

                            wbManufaDT.LoadCompleted += new LoadCompletedEventHandler(ManufaDT_LoadCompleted);
                        });
                        //2019-10-31 End add


                    });

                });

            }
        }

        private void wbBWNoHob_LoadedFrameComplete(object sender, Awesomium.Core.FrameEventArgs e)
        {
            isFirstNavigationBWNoHob = false;
            Awesomium.Windows.Controls.WebControl wb = sender as Awesomium.Windows.Controls.WebControl;
            dynamic document = (Awesomium.Core.JSObject)wb.ExecuteJavascriptWithResult("document");
            using (document)
            {
                if (isFirstLoadBWNoHob)
                {

                    //get user login info in App.config
                    string user = ConfigurationManager.AppSettings["user"];
                    string pass = ConfigurationManager.AppSettings["pass"];

                    //login
                    dynamic userInput = document.getElementById("txtUserId");
                    dynamic passInput = document.getElementById("txtPassword");
                    dynamic btnLogin = document.getElementById("cmdLogin");
                    userInput.setAttribute("value", user);
                    passInput.setAttribute("value", pass);
                    btnLogin.click();
                    isFirstLoadBWNoHob = false;
                    //điều hướng trang về 0 để vào trang in
                    pageNumberBWNoHob = 0;
                }
                #region lock
                else
                {
                    if (pageNumberBWNoHob == 0)
                    {
                        //nếu trang tải xong
                        if (wb.IsLoaded)
                        {

                            //check session time out
                            if (CheckSessionExpiredAwesomiumBrowser(wb))
                            {
                                //navigate to login page
                                isFirstLoadBWNoHob = true; //true: login page <> false: other pages     
                                document.getElementById("btnOK").click();
                            }
                            else
                            {
                                Awesomium.Core.JSObject window = wb.ExecuteJavascriptWithResult("window");
                                var navigate = wb.ExecuteJavascriptWithResult("window.location = '/manufa/PrintViewInstruction.aspx';");
                                //chuyển trang 
                                pageNumberBWNoHob = 1;
                            }
                        }

                    }
                    else if (pageNumberBWNoHob == 1)
                    {
                        //check session time out
                        if (CheckSessionExpiredAwesomiumBrowser(wb))
                        {
                            //2019-10-11 Start add: if BWNoHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                            _BWNoHob_IsOutOfSession = true;                            
                            //2019-10-11 End add

                            //navigate to login page
                            isFirstLoadBWNoHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {
                            //2019-10-11 Start add
                            if(_BWNoHob_IsOutOfSession)
                            {
                                //_BWNoHob_IsOutOfSession = false;//reset to default value

                                if (BWNoHob_IsPrinting) //nếu BWNoHob đang in
                                {
                                    //onclick btnClear to navigate to page 3
                                    pageNumberBWNoHob = 2;
                                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    btnClear.click();
                                }else if(BWF1_IsPrinting) //nếu máy BWF1 đang in
                                {
                                    pageNumberBWNoHob = 2;
                                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    btnClear.click();
                                }

                            }
                            //2019-10-11 End add
                            else
                            {
                                var printIndividual = wb.ExecuteJavascriptWithResult(@"                                                                              
                                        //set selected status option for the order that not printed yet
                                        var printMode = document.getElementById('ContentPlaceHolder1_rdoSelectOne');
                                        printMode.checked = true;");
                                if (printIndividual)
                                    pageNumberBWNoHob = 2;
                            }
                        }

                    }
                    else if (pageNumberBWNoHob == 2)
                    {
                        //2019-11-14 Start add: cập nhật lại trạng thái session sau khi re-login và reset trạng thái cho nút PRINT, thông báo in lại
                        if (_BWNoHob_IsOutOfSession)
                        {
                            _BWNoHob_IsOutOfSession = false;
                            IsGettingDataBW = true;
                            BWNoHob_IsGettingData = true;
                            BWNoHob_IsPrinting = false;

                            BWF1_IsGettingData = true;
                            BWF1_IsPrinting = false;

                            BW_NoHobStatus = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                            BW_F1Status = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                        }
                        //2019-11-14 End add

                    }
                    else if (pageNumberBWNoHob == 3)
                    {
                        if (CheckSessionExpiredAwesomiumBrowser(wb))
                        {
                            //2019-10-11 Start add: if BWNoHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                            _BWNoHob_IsOutOfSession = true;
                            //2019-10-11 End add

                            //thông báo đăng nhập lại khi hết Session
                            BW_NoHobStatus = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                            BW_F1Status = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                            //navigate to login page
                            isFirstLoadBWNoHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();
                        }
                        else
                        {
                            var print = wb.ExecuteJavascriptWithResult(@"                                                                              
                                         var orderToPrint = document.getElementById('ContentPlaceHolder1_grdPrintList_chkOutput_0');
                                        var btnPrint = document.getElementById('ContentPlaceHolder1_cmdPrint');

                                        if(orderToPrint != null)
                                        {   
                                            //check vào PO cần in
                                            orderToPrint.checked = 'checked';   
                                            //chọn máy in, 15 = pulley máy 2
                                            var printer = document.getElementById('ContentPlaceHolder1_drpPrinter');
                                            //2019-12-05 Manufa thay đổi vị trí máy in
                                            //printer.selectedIndex = 15
                                            printer.value = 121; 
                                            //máy pulley 2 value = 121
                                            //2019-12-05 
                                            //in 
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },500);

                                        }");
                            if (print)
                            {

                                pageNumberBWNoHob = 4; //chuyển trang sau khi in, lấy status
                            }


                        }


                    }
                    else if (pageNumberBWNoHob == 4)
                    {
                        if (wb.IsLoaded)
                        {
                            if (CheckSessionExpiredAwesomiumBrowser(wb))
                            {
                                //navigate to login page
                                isFirstLoadBWNoHob = true; //true: login page <> false: other pages     
                                document.getElementById("btnOK").click();
                            }
                            else
                            {
                                //lấy status sau khi in
                                string status = document.getElementById("ContentPlaceHolder1_AlertArea").innerHTML;
                                //nếu in thành công
                                if (status.ToLower().Contains("completed"))
                                {
                                    if (BWNoHob_IsPrinting) //nếu BWNoHob đang in
                                    {

                                        NoHobBW_TeethShape = SelectedOrderBWNoHob.TeethShape;
                                        NoHobBW_TeethQty = SelectedOrderBWNoHob.TeethQuantity;

                                        //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                        var poToDelete = ListOrdersBWNoHob.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWNoHob.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersBWNoHob.Remove(poToDelete);

                                            BW_NoHobStatus = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                            Task.Delay(1000).Wait();
                                            //2019-10-15 Start add: enable timer again
                                            timerBWHob.Start();
                                            //2019-10-15 End add
                                        }

                                        IsGettingDataBW = true;
                                        BWNoHob_IsGettingData = true;
                                        BWNoHob_IsPrinting = false;
                                    }else if(BWF1_IsPrinting) //nếu BWF1 đang in
                                    {
                                        F1BW_Diameter = SelectedOrderBWF1.Diameter_d;
                                        //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                        var poToDelete = ListOrdersBWF1.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWF1.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersBWF1.Remove(poToDelete);

                                            BW_F1Status = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                            Task.Delay(1000).Wait();
                                            //2019-10-15 Start add: enable timer again
                                            timerBWHob.Start();
                                            //2019-10-15 End add
                                        }

                                        IsGettingDataBW = true;
                                        BWF1_IsGettingData = true;
                                        BWF1_IsPrinting = false;
                                    }
                                }
                                else if (status.ToLower().Contains("found"))
                                {
                                    //PO chưa được chọn

                                    if (BWNoHob_IsPrinting) //nếu BWNoHob đang in
                                    {

                                        var poToDelete = ListOrdersBWNoHob.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWNoHob.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersBWNoHob.Remove(poToDelete);
                                            BW_NoHobStatus = "PO: " + poToDelete + " has been printed!"; ;
                                            Task.Delay(1000).Wait();
                                            //2019-10-15 Start add: enable timer again
                                            timerBWHob.Start();
                                            //2019-10-15 End add
                                        }

                                        BWNoHob_IsGettingData = true;
                                        BWNoHob_IsPrinting = false;
                                    }else if(BWF1_IsPrinting) //nếu BWF1 đang in
                                    {
                                        var poToDelete = ListOrdersBWF1.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWF1.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersBWF1.Remove(poToDelete);
                                            BW_F1Status = "PO: " + poToDelete + " has been printed!"; ;
                                            Task.Delay(1000).Wait();
                                            //2019-10-15 Start add: enable timer again
                                            timerBWHob.Start();
                                            //2019-10-15 End add
                                        }

                                        BWF1_IsGettingData = true;
                                        BWF1_IsPrinting = false;
                                    }

                                    //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    btnClear.click();
                                    //2019-10-04 End add

                                    //Enable btnGetData
                                    IsGettingDataBW = true;

                                    //2019-10-07 End add
                                    pageNumberBWNoHob = -1;//quay lại từ đầu
                                }
                                else
                                {
                                    //máy in hết giấy, kẹt giấy, offline
                                    //quay về trang in
                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable BWHobLoading if getting error
                                    IsGettingDataBW = true;
                                    if (BWNoHob_IsPrinting) //nếu BWNoHob đang in
                                    {
                                        //xóa PO đã in thành công khỏi list

                                        //2019-10-11 Start add: print message: Order has been printed
                                        BW_NoHobStatus = "Printer offline or paper ran out, please check it and try again!";
                                        //2019-10-07 End add
                                        BWNoHob_IsGettingData = true;
                                        BWNoHob_IsPrinting = false;
                                    }else if (BWF1_IsPrinting) //nếu BWF1 đang in
                                    {
                                        //xóa PO đã in thành công khỏi list

                                        //2019-10-11 Start add: print message: Order has been printed
                                        BW_F1Status = "Printer offline or paper ran out, please check it and try again!";
                                        //2019-10-07 End add
                                        BWF1_IsGettingData = true;
                                        BWF1_IsPrinting = false;
                                    }

                                    //2019-10-07 End add

                                    pageNumberBWNoHob = -1;
                                }

                            }
                            //2019-11-27 Start add: calculate total PO and total Orders after print BW
                            Calc_PO_and_Orders_BW();
                            //2019-11-27 End add
                        }
                    }
                    else if (pageNumberBWNoHob == 5)
                    {                       
                    }
                }
                #endregion
            }
        }

        //KEY-Manufa: Manufa_LoadCompleted
        private void ManufaDT_LoadCompleted(object sender, NavigationEventArgs e)
        {
            WebBrowser wb = sender as WebBrowser;
            mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;
            if (_mnfBWpage == 0)
            {
                //get user login info in App.config
                string user = ConfigurationManager.AppSettings["user"];
                string pass = ConfigurationManager.AppSettings["pass"];

                //login
                mshtml.IHTMLElement userInput = document.getElementById("txtUserId");
                mshtml.IHTMLElement passInput = document.getElementById("txtPassword");
                mshtml.IHTMLElement btnLogin = document.getElementById("cmdLogin");
                userInput.setAttribute("value", user);
                passInput.setAttribute("value", pass);
                btnLogin.click();
                _mnfBWpage = 1;
            }
            else if (_mnfBWpage == 1)
            {
                if (wb.IsLoaded)
                {
                    //http://10.4.24.111:8080/fp_mlt/PrintView.aspx
                    var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                    var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                    script.text = @" 
                                   window.onload = function(){
                                        //navigate 'Manufacturing Instructions Print' menu
                                        window.location = '/fp_mlt/PrintView.aspx';
                                   }
                              ";
                    head.appendChild((mshtml.IHTMLDOMNode)script);


                }
                _mnfBWpage = 2;
            }
            else if (_mnfBWpage == 2)
            {
                if (wb.IsLoaded)
                {
                    var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                    var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                    script.text = @" 
                                   window.onload = function(){
                                        var select = document.getElementById('ContentPlaceHolder1_drpStatus');
                                        select.selectedIndex = 2; //order not printed yet
                                        document.getElementById('ContentPlaceHolder1_txtLine').value = 'LE';

                                        var displayItem = document.getElementById('ContentPlaceHolder1_drpDispCount');                                        
                                        displayItem.selectedIndex = 1;//choose 500 items                                        

                                        var btnSearch = document.getElementById('ContentPlaceHolder1_cmbSelect');
                                        btnSearch.click();
                                                                               
                                   }
                              ";
                    head.appendChild((mshtml.IHTMLDOMNode)script);
                    HideScriptErrors(wb, true);


                }
                _mnfBWpage = 3;

            }
            else if (_mnfBWpage == 3)
            {
                if (wb.IsLoaded)
                {

                    //2019-10-31 Start add
                    //lấy html cả trang web
                    //var htmlPage = document.documentElement.outerHTML;
                    var htmltb = document.getElementById("ContentPlaceHolder1_sprPrintList_viewport").outerHTML;
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmltb);


                    //lấy raw manufa data
                    var rawManufaDT = GetRawManufaDT(doc);
                    //chuẩn bị data manufa cho BW
                    if (_mnfBW_getData)
                    {
                        PreparingBWManufaDT(rawManufaDT);
                    }

                    //chuẩn bị data manufa cho CNC
                    if (_mnfCNC_getData)
                    {
                        PreparingCNCManufaDT(rawManufaDT);
                    }
                    //2019-10-31 End add


                }
                _mnfBWpage = 0;//quay về
            }
            else if (_mnfBWpage == 4)
            {

                _mnfBWpage = 0;//quay về
            }
            else
            {

            }
        }



        //2019-10-31 Start add
        private List<OrderModel> GetRawManufaDT(HtmlAgilityPack.HtmlDocument doc)
        {

            List<OrderModel> lst = new List<OrderModel>();

            var rows = doc.DocumentNode.SelectNodes("//*[@id='ContentPlaceHolder1_sprPrintList_viewport']/tbody/tr");

            DateTime validDatetime;
            if (rows != null && rows.Count > 0)
            {
                for (int i = 0; i < rows.Count; i++)
                {

                    var cols = rows[i].SelectNodes("//*[@id='ContentPlaceHolder1_sprPrintList_viewport']/tbody/tr[" + (i + 1) + "]/td");
                    if (cols != null)
                    {
                        lst.Add(new OrderModel()
                        {
                            No = cols[0].InnerText.ToString() ?? "",
                            Sales_Order_No = cols[1].InnerText.ToString() ?? "",
                            Pier_Instruction_No = cols[2].InnerText.ToString() ?? "",
                            Manufa_Instruction_No = cols[3].InnerText.ToString() ?? "",
                            Global_Code = cols[4].InnerText.ToString() ?? "",
                            Customers = cols[5].InnerText.ToString() ?? "",
                            Item_Name = cols[6].InnerText.ToString() ?? "",
                            MC = cols[7].InnerText.ToString() ?? "",
                            //2019-10-03 Start delete
                            //Received_Date = !string.IsNullOrEmpty(worksheet.Cells[i, 9].Value.ToString()) ? Convert.ToDateTime(worksheet.Cells[i, 9].Value) : DateTime.Now,
                            //2019-10-03 End delete
                            Received_Date = DateTime.TryParse((string)cols[8].InnerText, out validDatetime) ? validDatetime : (DateTime?)null,
                            Factory_Ship_Date = DateTime.TryParse((string)cols[9].InnerText, out validDatetime) ? validDatetime : (DateTime?)null,
                            //mai chỉnh lại kiểu ngày mặc định
                            Number_of_Orders = cols[10].InnerText.ToString() ?? "",
                            Number_of_Available_Instructions = cols[11].InnerText.ToString() ?? "",
                            Number_of_Repairs = cols[12].InnerText.ToString() ?? "",
                            Number_of_Instructions = cols[13].InnerText.ToString() ?? "",
                            Line = cols[14].InnerText.ToString() ?? "",
                            PayWard = cols[15].InnerText.ToString() ?? "",
                            Major = cols[16].InnerText.ToString() ?? "",
                            Special_Orders = cols[17].InnerText.ToString() ?? "",
                            Method = cols[18].InnerText.ToString() ?? "",
                            Destination = cols[19].InnerText.ToString() ?? "",
                            Instructions_Print_date = DateTime.TryParse((string)cols[20].InnerText, out validDatetime) ? validDatetime : (DateTime?)null,
                            Latest_progress = cols[21].InnerText.ToString() ?? "",
                            Tack_Label_Output_Date = cols[22].InnerText.ToString() ?? "",
                            Completion_Instruction_Date = DateTime.TryParse((string)cols[23].InnerText, out validDatetime) ? validDatetime : (DateTime?)null,
                            Re_print_Count = cols[24].InnerText.ToString() ?? "",
                            Latest_issue_time = cols[25].InnerText.ToString() ?? "",

                            Material_code1 = cols[26].InnerText.ToString() ?? "",
                            Material_text1 = cols[27].InnerText.ToString() ?? "",
                            Amount_used1 = cols[28].InnerText.ToString() ?? "",
                            Unit1 = cols[29].InnerText.ToString() ?? "",
                            Material_code2 = cols[30].InnerText.ToString() ?? "",
                            Material_text2 = cols[31].InnerText.ToString() ?? "",
                            Amount_used2 = cols[32].InnerText.ToString() ?? "",
                            Unit2 = cols[33].InnerText.ToString() ?? "",
                            Material_code3 = cols[34].InnerText.ToString() ?? "",
                            Material_text3 = cols[35].InnerText.ToString() ?? "",
                            Amount_used3 = cols[36].InnerText.ToString() ?? "",
                            Unit3 = cols[37].InnerText.ToString() ?? "",
                            Material_code4 = cols[38].InnerText.ToString() ?? "",
                            Material_text4 = cols[39].InnerText.ToString() ?? "",
                            Amount_used4 = cols[40].InnerText.ToString() ?? "",
                            Unit4 = cols[41].InnerText.ToString() ?? "",
                            Material_code5 = cols[42].InnerText.ToString() ?? "",
                            Material_text5 = cols[43].InnerText.ToString() ?? "",
                            Amount_used5 = cols[44].InnerText.ToString() ?? "",
                            Unit5 = cols[45].InnerText.ToString() ?? "",
                            Material_code6 = cols[46].InnerText.ToString() ?? "",
                            Material_text6 = cols[47].InnerText.ToString() ?? "",
                            Amount_used6 = cols[48].InnerText.ToString() ?? "",
                            Unit6 = cols[49].InnerText.ToString() ?? "",
                            Material_code7 = cols[50].InnerText.ToString() ?? "",
                            Material_text7 = cols[51].InnerText.ToString() ?? "",
                            Amount_used7 = cols[52].InnerText.ToString() ?? "",
                            Unit7 = cols[53].InnerText.ToString() ?? "",
                            Material_code8 = cols[54].InnerText.ToString() ?? "",
                            Material_text8 = cols[55].InnerText.ToString() ?? "",
                            Amount_used8 = cols[56].InnerText.ToString() ?? "",
                            Unit8 = cols[57].InnerText.ToString() ?? "",
                            Inner_Code = BWCNCViewModel_Sub.Instance.GetF1OrdersInnerCode(cols[58].InnerText.ToString()) ?? "",
                            Classify_Code = cols[59].InnerText.ToString() ?? "",
                        });
                    }

                }
            }


            return lst;

        }

        private async void PreparingBWManufaDT(List<OrderModel> lstManufaInfo)
        {

            //get InnerCode from lstManufaInfo 
            //2020-04-16 Start edit: thêm function tách Product cho đơn F1
            var innerCodes = string.Format("'{0}'", string.Join("','", lstManufaInfo.Select(c => c.Inner_Code)));
            //var innerCodes = string.Format("'{0}'", string.Join("','", lstManufaInfo.Select(c => BWCNCViewModel_Sub.Instance.GetF1OrdersInnerCode(c.Inner_Code))));
            //2020-01-16 End edit
            List<OrderModel> lstPulleyMaster = OrderBLL.Instance.GetPulleyMasterByProduct(innerCodes);
            //get LineAB
            //2019-10-05 Start add: get machine types to identify which PO belongs to BW or CNC
            List<MachineTypeModel> lstMachineTypes = await MachineTypeBLL.Instance.GetMachineTypes();
            //2019-10-05 End add
          
            if (lstPulleyMaster.Count > 0)
            {


                IEnumerable<OrderModel> result = (from r1 in lstManufaInfo
                                                  join r2 in lstPulleyMaster on r1.Inner_Code equals r2.Inner_Code
                                                  join r3 in lstMachineTypes on r1.Material_code1 equals r3.MaterialCode
                                                  select new OrderModel()
                                                  {
                                                      //No = r1.No,
                                                      Sales_Order_No = r1.Sales_Order_No,
                                                      Pier_Instruction_No = r1.Pier_Instruction_No,
                                                      Manufa_Instruction_No = r1.Manufa_Instruction_No,
                                                      Global_Code = r1.Global_Code,
                                                      //Customers = r1.Customers,
                                                      Item_Name = r1.Item_Name,
                                                      MC = r1.MC,
                                                      Received_Date = r1.Received_Date,
                                                      Factory_Ship_Date = r1.Factory_Ship_Date,
                                                      Number_of_Orders = r1.Number_of_Orders,
                                                      Number_of_Available_Instructions = r1.Number_of_Available_Instructions,
                                                      //Number_of_Repairs = r1.Number_of_Repairs,
                                                      //Number_of_Instructions = r1.Number_of_Instructions,
                                                      Line = r1.Line,
                                                      //PayWard = r1.PayWard,
                                                      //Major = r1.Major,
                                                      //Special_Orders = r1.Special_Orders,
                                                      //Method = r1.Method,
                                                      //Destination = r1.Destination,
                                                      //Instructions_Print_date = r1.Instructions_Print_date,
                                                      //Latest_progress = r1.Latest_progress,
                                                      //Tack_Label_Output_Date = r1.Tack_Label_Output_Date,
                                                      //Completion_Instruction_Date = r1.Completion_Instruction_Date,
                                                      //Re_print_Count = r1.Re_print_Count,
                                                      //Latest_issue_time = r1.Latest_issue_time,
                                                      Material_code1 = r1.Material_code1,
                                                      Material_text1 = r1.Material_text1.Substring(r1.Material_text1.LastIndexOf("/") + 1, 3),
                                                      Amount_used1 = r1.Amount_used1,
                                                      Unit1 = r1.Unit1,
                                                      //Material_code2 = r1.Material_code2,
                                                      //Material_text2 = r1.Material_text2,
                                                      //Amount_used2 = r1.Amount_used2,
                                                      //Unit2 = r1.Unit2,
                                                      //Material_code3 = r1.Material_code3,
                                                      //Material_text3 = r1.Material_text3,
                                                      //Amount_used3 = r1.Amount_used3,
                                                      //Unit3 = r1.Unit3,
                                                      //Material_code4 = r1.Material_code4,
                                                      //Material_text4 = r1.Material_text4,
                                                      //Amount_used4 = r1.Amount_used4,
                                                      //Unit4 = r1.Unit4,
                                                      //Material_code5 = r1.Material_code5,
                                                      //Material_text5 = r1.Material_text5,
                                                      //Amount_used5 = r1.Amount_used5,
                                                      //Unit5 = r1.Unit5,
                                                      //Material_code6 = r1.Material_code6,
                                                      //Material_text6 = r1.Material_text6,
                                                      //Amount_used6 = r1.Amount_used6,
                                                      //Unit6 = r1.Unit6,
                                                      //Material_code7 = r1.Material_code7,
                                                      //Material_text7 = r1.Material_text7,
                                                      //Amount_used7 = r1.Amount_used7,
                                                      //Unit7 = r1.Unit7,
                                                      //Material_code8 = r1.Material_code8,
                                                      //Material_text8 = r1.Material_text8,
                                                      //Amount_used8 = r1.Amount_used8,
                                                      //Unit8 = r1.Unit8,
                                                      Inner_Code = r1.Inner_Code,
                                                      //Classify_Code = r1.Classify_Code,
                                                      Hobbing = r2.Hobbing,
                                                      TeethQuantity = r2.TeethQuantity,
                                                      TeethShape = r2.Hobbing ? GetKnifeType(GetTeethPart(r2.TeethShape), r2.TeethQuantity) : "",
                                                      LineAB = r3.LineAB,
                                                      LineName = r3.LineName,
                                                      Diameter_d = GetDiameter(r2.Diameter_d, r1.Item_Name)
                                                  }).ToList();

                BWOrders = new ObservableCollection<OrderModel>(result);
                //lấy danh sách BW orders
                var _BWOrders = BWOrders.Where(c => c.Hobbing && c.LineAB.Equals("BW"));

                //2020-03-20 Start add: lấy ra các orders là F1
                var ordersF1 = _BWOrders.Where(c => BWCNCViewModel_Sub.Instance.IsF1Order(c.Item_Name)).ToList();
                ListOrdersBWF1 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.Sort_F1Orders(ordersF1,F1BW_Diameter));

                //loại các order f1 khỏi _BWOrders
                _BWOrders = _BWOrders.Except(ordersF1);
                //2020-03-20 End add

                //lấy ra các orders có TeethShape = Hob11BW TeethShape
                var ordersBW_11 = _BWOrders.Where(c => c.TeethShape.Equals(Hob11BW_TeethShape)).ToList();
                //chia đôi số lượng items của BWOrders
                double chunkBW = Math.Ceiling((_BWOrders.GroupBy(g => g.TeethShape).Count() * 1.0) / 2);

                //2019-10-29 Start add       
                //nếu BWHob11 có TeethShape và TeethQty giống
                if (ordersBW_11.Count > 0)
                {
                    //nếu có sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in                               
                    //var sortedList = SortOrders(ordersBW_11, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter); //old way
                    var sortedList = SetupSSD_BW(ordersBW_11, SSDType.N1,Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter); //new way
                    ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedList);

                    //2019-11-26 Start add:
                    //lấy ra phần còn lại kiểm tra nếu BWHob12 có teethShape giống thì chia cho BWHob12
                    var restOfListBWHob11 = _BWOrders.Except(ListOrdersBWHob11).ToList();
                    if (restOfListBWHob11.Count > 0)
                    {
                        //lấy hết phần còn lại của restOfListBWHob11 đẩy về cho BWHob12

                        //08-10-2020: tạm thời chuyển qua ưu tiên SSD N+1 trước
                        //var lstSorted = SortOrders(restOfListBWHob11, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter); //old way
                        var lstSorted = SetupSSD_BW(restOfListBWHob11, SSDType.N1,Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter); //new way
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(lstSorted);
                       
                    }
                    //2019-11-26 End add
                }
                else
                {
                    
                    //2019-11-26 Start add:
                    //nếu ko kiểm tra BWHob12 có TeethShape và TeethQty giống hay ko
                    //lấy ra các orders có TeethShape = Hob12BW TeethShape
                    var _BWHob12orders = _BWOrders.Where(c => c.TeethShape.Equals(Hob12BW_TeethShape)).ToList();
                    if (_BWHob12orders.Count > 0)
                    {
                        //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                        //var _ordersBWHob12 = SortOrders(_BWHob12orders, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter);//old way
                        var _ordersBWHob12 = SetupSSD_BW(_BWHob12orders, SSDType.N1,Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter);//new way
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(_ordersBWHob12);

                        //phần còn lại sau khi chia cho BWHob12 sẽ đẩy qua BWHob11
                        var theRestOfBWHob12 = _BWOrders.Except(ListOrdersBWHob12).ToList();
                        if (theRestOfBWHob12.Count > 0)
                        {
                            //var _sorted_theRestOfBWHob12 = SortOrders(theRestOfBWHob12, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter);//old way
                            var _sorted_theRestOfBWHob12 = SetupSSD_BW(theRestOfBWHob12, SSDType.N1,Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter);//new way
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(_sorted_theRestOfBWHob12);

                        }
                    }
                    else
                    {
                        //nếu ko chia đôi số lượng items của BWOrders                              
                        var tempLst = new List<OrderModel>();
                        _BWOrders.GroupBy(g => g.TeethShape).Take((int)chunkBW).Select(g => g.ToList()).ToList().ForEach(i => tempLst.AddRange(i));
                        //var sortedList = SortOrders(tempLst, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter);//old way
                        var sortedList = SetupSSD_BW(tempLst, SSDType.N1,Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter);//new way
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedList);

                        //lấy ra các orders có TeethShape = Hob12 BW và không nằm trong ListOrdersBWHob11
                        _BWOrders.Where(c => ListOrdersBWHob11.All(c2 => !c2.TeethShape.Equals(c.TeethShape))).ToList();
                        var ordersBW_12 = _BWOrders.Except(ListOrdersBWHob11).ToList();

                        //2019-10-29 Start add
                        if (ordersBW_12.Count > 0)
                        {
                            //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                            //var sortedList2 = SortOrders(ordersBW_12, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter);//old way
                            var sortedList2 = SetupSSD_BW(ordersBW_12, SSDType.N1,Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter);//new way
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(sortedList2);

                        }
                        //2019-10-29 End add
                    }
                    //2019-11-26 End add
                }

                //2019-10-29 End add

               
                //BW orders with No Hob                       
                ListOrdersBWNoHob = new ObservableCollection<OrderModel>(BWOrders.Where(c => c.Hobbing == false && c.LineAB.Equals("BW")).ToList().OrderByDescending(c => NoHobBW_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));


                //2019-12-02 Start change: đổi cách đọc thông tin PO từ JSON để đồng bộ BWCNC và CNC từ MYSQL
                //BWCNC Hob9       
                //đọc PO đã delete của Hob9 sau khi in và lấy ra TeethShape, TeethQuantity cho việc sắp xếp
                //OrderModel deletedPO_Hob9 = JsonHelper.ReadJson2<OrderModel>(@"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob9.json");
                try
                {
                    OrderModel deletedPO_Hob9 = new OrderModel(MySQLProvider.Instance.ExecuteQuery("SELECT * FROM CNCHob9").Rows[0]);
                    OrderModel deletedPO_Hob10 = new OrderModel(MySQLProvider.Instance.ExecuteQuery("SELECT * FROM CNCHob10").Rows[0]);
                    //lấy danh sách các orders của CNC
                    var CNCOrders = BWOrders.Where(c => c.Hobbing && c.LineAB.Equals("CNC"));

                    //2019-11-30 Start add: 
                    //B1: tách các orders còn lại sau khi chia cho BWHob11 và BWHob12
                    if (ListOrdersBWHob11 != null)
                    {
                        var restOfBWOrders = _BWOrders.Except(ListOrdersBWHob11).ToList();
                        if (ListOrdersBWHob12 != null)
                        {
                            //2019-12-24 Start edit: thiếu restOfBWOrders = ...
                            restOfBWOrders = restOfBWOrders.Except(ListOrdersBWHob12).ToList();
                            //2019-12-24 End edit
                        }
                        //B2: kiểm tra BWCNCHob9_HasBWOrders và BWCNCHob10_HasOrders

                        //Nếu BWCNCHob9_HasBWOrders
                        //2020-02-08 Start edit
                        var bw_cnc9_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob9();
                        var bw_cnc10_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob10();
                        //2020-02-08 End edit
                        if (bw_cnc9_HasBWOrders == 1)
                        {
                            //nếu BWCNCHob9_HasOrders => tách các orders trong restOfBWOrders rồi tìm các item có TeethShape, TeethQty = Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty
                            var _new_BW_OrdersForCNC9 = restOfBWOrders.Where(c => c.TeethShape.Equals(Hob9BWCNC_TeethShape)).ToList();
                            //nếu _new_BW_OrdersForCNC9 có items
                            if (_new_BW_OrdersForCNC9 != null && _new_BW_OrdersForCNC9.Count > 0)
                            {

                                //sắp xếp các ordersBW trong _new_BW_OrdersForCNC9 theo thứ tự TeethShape, TeethQty, Diameter_d của BWCNCHob9

                                //var _sorted_BW_Orders = SortOrders(_new_BW_OrdersForCNC9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                                var _sorted_BW_Orders = SetupSSD_BW(_new_BW_OrdersForCNC9, SSDType.N1,Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);

                                //gán orders BW được sắp xếp vào ListOrdersBWCNCHob9
                                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(_sorted_BW_Orders);
                                if (CNCOrders != null && CNCOrders.Count() > 0)
                                {
                                    //tìm các đơn CNC có TeethShape giống với TeethShape của item cuối cùng trong _sorted_BW_Orders
                                    var _last_item_of_sorted_BW_Orders = _sorted_BW_Orders.LastOrDefault();
                                    if (_last_item_of_sorted_BW_Orders != null)
                                    {
                                        var _new_CNC_OrdersForCNC9 = CNCOrders.Where(c => c.TeethShape.Equals(_last_item_of_sorted_BW_Orders.TeethShape)).ToList();

                                        //sắp xếp list _new_CNC_OrdersForCNC9 theo thứ tự TeethShape, TeethQty, Diameter_d của _last_item_of_sorted_BW_Orders
                                        //var _sorted_CNC_OrdersForCNC9 = SortOrders(_new_CNC_OrdersForCNC9, _last_item_of_sorted_BW_Orders.TeethShape, _last_item_of_sorted_BW_Orders.TeethQuantity, _last_item_of_sorted_BW_Orders.Diameter_d);
                                        var _sorted_CNC_OrdersForCNC9 = SetupSSD_CNC(_new_CNC_OrdersForCNC9, SSDType.N1,Hob9BWCNC_DMat,_last_item_of_sorted_BW_Orders.TeethShape, _last_item_of_sorted_BW_Orders.TeethQuantity, _last_item_of_sorted_BW_Orders.Diameter_d);

                                        //add thêm các items của _sorted_CNC_OrdersForCNC9 vào  ListOrdersBWCNCHob9 (nối tiếp sau item BW cuối cùng)
                                        _sorted_CNC_OrdersForCNC9.ForEach(c => ListOrdersBWCNCHob9.Add(c));
                                    }
                                }

                            }
                            else
                            {
                                //26-10-2020: nếu ko có đơn BW cho CNC9 thì cập nhật trạng thái lock máy CNC9 thành 0
                                BLL.OrderBLL.Instance.SetLockCNC9(0);
                                //26-10-2020
                            }
                        }

                        if (bw_cnc10_HasBWOrders == 1) //Nếu BWCNCHob10_HasBWOrders
                        {
                            //nếu BWCNCHob10_HasOrders => tách các orders trong restOfBWOrders (trừ đi các items của ListOrdersBWCNCHob9) rồi tìm các item có TeethShape, TeethQty = Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty
                            var _new_BW_OrdersForCNC10 = restOfBWOrders.Except(ListOrdersBWCNCHob9).Where(c => c.TeethShape.Equals(Hob10BWCNC_TeethShape)).ToList();
                            //nếu _new_BW_OrdersForCNC10 có items
                            if (_new_BW_OrdersForCNC10 != null && _new_BW_OrdersForCNC10.Count > 0)
                            {
                                //sắp xếp các ordersBW trong _new_BW_OrdersForCNC10 theo thứ tự TeethShape, TeethQty, Diameter_d của BWCNCHob10

                                //var _sorted_BW_Orders = SortOrders(_new_BW_OrdersForCNC10, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                                var _sorted_BW_Orders = SetupSSD_BW(_new_BW_OrdersForCNC10, SSDType.N1,Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);

                                //gán orders BW được sắp xếp vào ListOrdersBWCNCHob9
                                ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(_sorted_BW_Orders);
                                if (CNCOrders.Except(ListOrdersBWCNCHob9) != null && CNCOrders.Except(ListOrdersBWCNCHob9).Count() > 0)
                                {
                                    //tìm các đơn CNC còn lại (sau khi bỏ đi các đơn CNC đã có trong BWCNCHob9) có TeethShape giống với TeethShape của item cuối cùng trong _sorted_BW_Orders
                                    var _last_item_of_sorted_BW_Orders = _sorted_BW_Orders.LastOrDefault();
                                    if (_last_item_of_sorted_BW_Orders != null)
                                    {
                                        var _new_CNC_OrdersForCNC10 = CNCOrders.Except(ListOrdersBWCNCHob9).Where(c => c.TeethShape.Equals(_last_item_of_sorted_BW_Orders.TeethShape)).ToList();

                                        //sắp xếp list _new_CNC_OrdersForCNC10 theo thứ tự TeethShape, TeethQty, Diameter_d của _last_item_of_sorted_BW_Orders
                                        //var _sorted_CNC_OrdersForCNC10 = SortOrders(_new_CNC_OrdersForCNC10, _last_item_of_sorted_BW_Orders.TeethShape, _last_item_of_sorted_BW_Orders.TeethQuantity, _last_item_of_sorted_BW_Orders.Diameter_d);
                                        var _sorted_CNC_OrdersForCNC10 = SetupSSD_CNC(_new_CNC_OrdersForCNC10,SSDType.N1,Hob10BWCNC_DMat, _last_item_of_sorted_BW_Orders.TeethShape, _last_item_of_sorted_BW_Orders.TeethQuantity, _last_item_of_sorted_BW_Orders.Diameter_d);
                                        //add thêm các items của _sorted_CNC_OrdersForCNC10 vào ListOrdersBWCNCHob10 (nối tiếp sau item BW cuối cùng)
                                        _sorted_CNC_OrdersForCNC10.ForEach(c => ListOrdersBWCNCHob10.Add(c));
                                    }
                                }
                            }
                            else
                            {
                                //26-10-2020: nếu ko có đơn BW cho CNC10 thì cập nhật trạng thái lock máy CNC9 thành 0
                                BLL.OrderBLL.Instance.SetLockCNC9(0);
                                //26-10-2020
                            }

                        }
                        
                        if(bw_cnc9_HasBWOrders != 1 && bw_cnc10_HasBWOrders != 1)
                        {
                            //lấy ra các orders CNC có TeethShape = deletedPO_Hob9 TeethShape
                            var ordersBWCNC_9 = CNCOrders.Where(c => c.TeethShape.Equals(deletedPO_Hob9.TeethShape)).ToList();
                            //chia đôi số lượng items của CNCOrders
                            double chunkCNC = Math.Ceiling((CNCOrders.GroupBy(g => g.TeethShape).Count() * 1.0) / 2);
                            //nếu deletedPO_Hob9 có value
                            if (deletedPO_Hob9 != null)
                            {
                                //nếu có orders cùng TeethShape với CNCHob9
                                if (ordersBWCNC_9.Count > 0)
                                {
                                    //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                    //nếu có sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in                           
                                    //var sortedList = SortOrders_CNC(ordersBWCNC_9, deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d);
                                    var sortedList = SetupSSD_CNC(ordersBWCNC_9, SSDType.N1,deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d);

                                    //2020-01-04 End edit
                                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(sortedList);

                                    //2019-11-30 Start add: các orders CNC còn lại sau khi đổ vào BWCNCHob9 so sánh với TeethShape của CNCHob10
                                    var _rest_BWCNCHob9_Orders = CNCOrders.Except(ListOrdersBWCNCHob9).ToList();
                                    if (_rest_BWCNCHob9_Orders != null && _rest_BWCNCHob9_Orders.Count > 0)
                                    {
                                        //tìm các orders trong _rest_BWCNCHob9_Orders có TeethShape = TeethShape của CNCHob10
                                        if (deletedPO_Hob10 != null)
                                        {
                                            var _new_orders_BWCNCHob10 = _rest_BWCNCHob9_Orders.Where(c => c.TeethShape.Equals(deletedPO_Hob10.TeethShape)).ToList();
                                            if (_new_orders_BWCNCHob10.Count > 0)
                                            {
                                                //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                                //sắp xếp các items trong _new_orders_BWCNCHob10 theo thứ tự TeethShape, TeethQty, Diameter_d của deletedPO_Hob10
                                                //var _sorted_orders_BWCNCHob10 = SortOrders_CNC(_new_orders_BWCNCHob10, deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                                var _sorted_orders_BWCNCHob10 = SetupSSD_CNC(_new_orders_BWCNCHob10, SSDType.N1,deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                                //2020-01-04 End edit
                                                //gán _sorted_orders_BWCNCHob10 cho ListOrdersBWCNCHob10
                                                ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(_sorted_orders_BWCNCHob10);

                                            }
                                            else
                                            {
                                                //nếu ko tìm thấy orders CNC nào trong _rest_BWCNCHob9_Orders có TeethShape = TeethShape của CNCHob10
                                                //sắp xếp tất cả các orders CNC còn lại cho BWCNCHob10
                                                //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                                //var _sorted_orders_BWCNCHob10 = SortOrders_CNC(_rest_BWCNCHob9_Orders, deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                                var _sorted_orders_BWCNCHob10 = SetupSSD_CNC(_rest_BWCNCHob9_Orders,SSDType.N1, deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                                //2020-01-04 End edit
                                                //gán _sorted_orders_BWCNCHob10 cho ListOrdersBWCNCHob10
                                                ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(_sorted_orders_BWCNCHob10);

                                            }
                                        }

                                    }
                                    //2019-11-30 End add
                                }
                                else
                                {
                                    //2019-11-30 Start add
                                    //nếu ko tìm các orders có cùng TeethShape với CNCHob10
                                    if (deletedPO_Hob10 != null)
                                    {
                                        var _orders_for_BWCNCHob10 = CNCOrders.Where(c => c.TeethShape.Equals(deletedPO_Hob10.TeethShape)).ToList();
                                        if (_orders_for_BWCNCHob10.Count > 0)
                                        {
                                            //nếu _orders_for_BWCNCHob10 có items cùng TeethShape với CNCHob10
                                            //sắp xếp các items theo thứ tự TeethShape, TeethQty, Diameter_d của deletedPO_Hob10
                                            //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                            //var _sorted_orders_for_BWCNCHob10 = SortOrders_CNC(_orders_for_BWCNCHob10, deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                            var _sorted_orders_for_BWCNCHob10 = SetupSSD_CNC(_orders_for_BWCNCHob10,SSDType.N1, deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                            //2020-01-04 End edit
                                            //gán _sorted_orders_for_BWCNCHob10 cho ListOrdersBWCNCHob10
                                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(_sorted_orders_for_BWCNCHob10);

                                            //các orders CNC còn lại sau khi chia cho BWCNCHob10 được chia cho BWCNCHob9
                                            var _rest_BWCNCHob10_orders = CNCOrders.Except(ListOrdersBWCNCHob10).ToList();
                                            if (_rest_BWCNCHob10_orders != null && _rest_BWCNCHob10_orders.Count > 0)
                                            {
                                                //tìm các orders CNC trong _rest_BWCNCHob10_orders có TeethShape = TeethShape của CNCHob9
                                                if (deletedPO_Hob9 != null)
                                                {
                                                    var _new_orders_BWCNCHob9 = _rest_BWCNCHob10_orders.Where(c => c.TeethShape.Equals(deletedPO_Hob9.TeethShape)).ToList();
                                                    if (_new_orders_BWCNCHob9.Count > 0)
                                                    {
                                                        //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                                        //sắp xếp các items trong _new_orders_BWCNCHob9 theo thứ tự TeethShape, TeethQty, Diameter_d của deletedPO_Hob9
                                                        //var _sorted_orders_BWCNCHob9 = SortOrders_CNC(_new_orders_BWCNCHob9, deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d).ToList();
                                                        var _sorted_orders_BWCNCHob9 = SetupSSD_CNC(_new_orders_BWCNCHob9,SSDType.N1, deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d).ToList();
                                                        //2020-01-04 End edit
                                                        //gán _new_orders_BWCNCHob9 cho ListOrdersBWCNCHob9
                                                        ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(_sorted_orders_BWCNCHob9);

                                                    }
                                                    else
                                                    {
                                                        //nếu ko tìm thấy items nào trong _rest_BWCNCHob10_orders có TeethShape = TeethShape của CNCHob9
                                                        //sắp xếp tất cả các orders CNC còn lại cho BWCNCHob9
                                                        //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                                        //var _sorted_orders_BWCNCHob9 = SortOrders_CNC(_rest_BWCNCHob10_orders, deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d).ToList();
                                                        var _sorted_orders_BWCNCHob9 = SetupSSD_CNC(_rest_BWCNCHob10_orders,SSDType.N1, deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d).ToList();
                                                        //2020-01-04 End edit
                                                        //gán _sorted_orders_BWCNCHob9 cho ListOrdersBWCNCHob9
                                                        ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(_sorted_orders_BWCNCHob9);

                                                    }

                                                }
                                            }
                                        }
                                        else
                                        {
                                            //2019-11-30 Start lock
                                            //nếu ko tìm được các orders có cùng TeethShape với CNCHob9 và CNCHob10
                                            //nếu ko chia đôi số lượng items của CNCOrders            
                                            var tempLst = new List<OrderModel>();
                                            CNCOrders.GroupBy(g => g.TeethShape).Take((int)chunkCNC).Select(g => g.ToList()).ToList().ForEach(i => tempLst.AddRange(i));
                                            //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                            //var sortedList = SortOrders_CNC(tempLst, deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d);
                                            var sortedList = SetupSSD_CNC(tempLst,SSDType.N1, deletedPO_Hob9.Material_text1, deletedPO_Hob9.TeethShape, deletedPO_Hob9.TeethQuantity, deletedPO_Hob9.Diameter_d);
                                            //2020-01-04 End edit
                                            ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(sortedList);

                                            //các items còn lại chuyển qua BWCNCHob10
                                            var _rest_of_items_from_ListOrdersBWCNCHob9 = CNCOrders.Except(ListOrdersBWCNCHob9).ToList();
                                            //2020-01-04 Start edit: thay đổi SortOrders_CNC để ưu tiên đường kính vật liệu cho máy CNC
                                            //var _sorted_items_for_BWCNCHob10 = SortOrders_CNC(_rest_of_items_from_ListOrdersBWCNCHob9, deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                            var _sorted_items_for_BWCNCHob10 = SetupSSD_CNC(_rest_of_items_from_ListOrdersBWCNCHob9,SSDType.N1, deletedPO_Hob10.Material_text1, deletedPO_Hob10.TeethShape, deletedPO_Hob10.TeethQuantity, deletedPO_Hob10.Diameter_d).ToList();
                                            //2020-01-04 End edit
                                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(_sorted_items_for_BWCNCHob10);

                                            //2019-11-30 End lock
                                        }
                                    }

                                    //2019-11-30 End add

                                }
                            }

                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Không đọc được thông tin từ PO bên CNC, vui lòng kiểm tra lại đường truyền", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                
                //2019-12-02 End lock

            }

            //2020-01-04 Start lock: kiểm tra số lượng đơn
            //if (totalOrders != lstManufaInfo.Count)
            //{
               

            //}
            //else
            //{
            //    //nếu totalOrders == lstManufaInfo, ko làm gì hết.
                
            //}
            //20-01-04 End lock
            //2019-12-10 End add

                      
            //2019-12-16 Start add: sắp xếp lại các orders BW11 ưu tiên N, N+1
            var t = Convert.ToInt32(DateTime.Now.ToString("HH"));
                     
            try
            {
                int status = OrderBLL.Instance.Get_bwIsSetupN1();
                int dayAddToCheckN1 = HolidayHelper.Instance.DayCountN1; //DateTime.Now.ToString("dddd").Equals("Saturday") ? 2 : 1;
                //status 0: check SSD false => ko ưu tiên N, N+1, status 1: xếp N, N+1
                //thời gian từ 6:00 đến 23:59 => lấy đơn N+1
               
                if (status == 1 && (t>=6 && t<24))
                {
                    //lấy hết đơn N+1 của các máy (nếu thứ 7 thì +2)
                    var ordersN1BW11 = ListOrdersBWHob11 != null ? ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList() : new List<OrderModel>();
                    var ordersN1BW12 = ListOrdersBWHob12 != null ? ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList() : new List<OrderModel>();
                    var ordersN1BWCNC9 = ListOrdersBWCNCHob9 != null ? ListOrdersBWCNCHob9.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList() : new List<OrderModel>();
                    var ordersN1BWCNC10 = ListOrdersBWCNCHob10 != null ? ListOrdersBWCNCHob10.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList() : new List<OrderModel>();
                    var tempList = new List<OrderModel>();                 
                    tempList.AddRange(ordersN1BW11);                  
                    tempList.AddRange(ordersN1BW12);               
                    tempList.AddRange(ordersN1BWCNC9);               
                    tempList.AddRange(ordersN1BWCNC10);
                    if(tempList.Count > 0)
                        SetupOrdersBW_N1(tempList, "N+1");                 
                }
                else if(status == 1 && (t>=0 && t<6))
                {
                    //lấy hết đơn N các máy lst.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(1).Date).ToList();
                    var ordersNBW11 = ListOrdersBWHob11 != null ? ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList() : new List<OrderModel>();
                    var ordersNBW12 = ListOrdersBWHob12 != null ? ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList(): new List<OrderModel>();
                    var ordersNBWCNC9 = ListOrdersBWCNCHob9 != null ? ListOrdersBWCNCHob9.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList(): new List<OrderModel>();
                    var ordersNBWCNC10 = ListOrdersBWCNCHob10 != null ? ListOrdersBWCNCHob10.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList() : new List<OrderModel>();
                    var tempList = new List<OrderModel>();
                    tempList.AddRange(ordersNBW11);
                    tempList.AddRange(ordersNBW12);
                    tempList.AddRange(ordersNBWCNC9);
                    tempList.AddRange(ordersNBWCNC10);
                    if (tempList.Count > 0)
                        SetupOrdersBW_N1(tempList, "N");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }


            //2019-12-16 End add

        


            //2019-11-27 Start add: tính total PO và total Orders cho BW
            Calc_PO_and_Orders_BW();
            //2019-11-27 End add

            //2019-12-11 Start add: khi get Orders từ Manufa, enable tất cả các nút in
            BWHob11_IsGettingData = true;
            BWHob12_IsGettingData = true;
            BWNoHob_IsGettingData = true;
            BWCNCHob9_IsGettingData = true;
            BWCNCHob10_IsGettingData = true;
            //2019-12-11 End add

            //lấy manufa data cho BW xong
            _mnfBW_getData = false;
            //all task finished
            IsLoadingBW = false;
            IsGettingDataBW = true;
            timerBWHob.Start();
            //2019-11-14 Start add: sau khi lấy xong data cho BW thì enable timerGetData để mỗi giờ tự động lấy tiếp data
            timerGetData.Start();
            //2019-11-14 End add
        }



        private async void PreparingCNCManufaDT(List<OrderModel> lstManufaInfo)
        {
            await Task.Run(async () =>
            {

                try
                {
                    //get InnerCode from lstManufaInfo 
                    var innerCodes = string.Format("'{0}'", string.Join("','", lstManufaInfo.Select(c => c.Inner_Code)));
                    List<OrderModel> lstPulleyMaster = OrderBLL.Instance.GetPulleyMasterByProduct(innerCodes);
                    //get LineAB
                    //2019-10-05 Start add: get machine types to identify which PO belongs to BW or CNC
                    List<MachineTypeModel> lstMachineTypes = await MachineTypeBLL.Instance.GetMachineTypes();
                    //2019-10-05 End add

                    //2019-12-10 Start add: kiểm tra số lượng order lấy về từ Manufa: nếu lớn hơn hoặc nhỏ hơn số lượng các Orders của BW và CNC thì đổ dữ liệu mới
                    
                    //var count_OrdersCNC_Hob9 = ListOrdersCNCHob9 != null ? ListOrdersCNCHob9.Count : 0;
                    //var count_OrdersCNC_Hob10 = ListOrdersCNCHob10 != null ? ListOrdersCNCHob10.Count : 0;
                    //var count_OrdersCNC_NoHob = ListOrdersCNCNoHob != null ? ListOrdersCNCNoHob.Count : 0;
                    ////count từ màn hình CNC, Hob9,10 và NoHob
                    //var count_OrdersBW = OrderBLL.Instance.GetTotalBWOrders();

                    //var totalOrders = count_OrdersCNC_Hob9 + count_OrdersCNC_Hob10 + count_OrdersCNC_NoHob + count_OrdersBW;

                    if (lstPulleyMaster.Count > 0)
                    {


                        IEnumerable<OrderModel> result = (from r1 in lstManufaInfo
                                                          join r2 in lstPulleyMaster on r1.Inner_Code equals r2.Inner_Code
                                                          join r3 in lstMachineTypes on r1.Material_code1 equals r3.MaterialCode
                                                          select new OrderModel()
                                                          {
                                                              //No = r1.No,
                                                              Sales_Order_No = r1.Sales_Order_No,
                                                              Pier_Instruction_No = r1.Pier_Instruction_No,
                                                              Manufa_Instruction_No = r1.Manufa_Instruction_No,
                                                              Global_Code = r1.Global_Code,
                                                              //Customers = r1.Customers,
                                                              Item_Name = r1.Item_Name,
                                                              MC = r1.MC,
                                                              Received_Date = r1.Received_Date,
                                                              Factory_Ship_Date = r1.Factory_Ship_Date,
                                                              Number_of_Orders = r1.Number_of_Orders,
                                                              Number_of_Available_Instructions = r1.Number_of_Available_Instructions,                                                     
                                                              Line = r1.Line,                                                       
                                                              Material_code1 = r1.Material_code1,
                                                              Material_text1 = r1.Material_text1.Substring(r1.Material_text1.LastIndexOf("/") + 1, 3),
                                                              Amount_used1 = r1.Amount_used1,
                                                              Unit1 = r1.Unit1,                                                            
                                                              Inner_Code = r1.Inner_Code,                                                             
                                                              Hobbing = r2.Hobbing,
                                                              TeethQuantity = r2.TeethQuantity,
                                                              TeethShape = r2.Hobbing ? GetKnifeType(GetTeethPart(r2.TeethShape), r2.TeethQuantity) : "",
                                                              LineAB = r3.LineAB,
                                                              LineName = r3.LineName,
                                                              Diameter_d = GetDiameter(r2.Diameter_d, r1.Item_Name)
                                                          }).ToList();

                        CNCOrders = new ObservableCollection<OrderModel>(result);
                        var _CNCOrdersHob = CNCOrders.Where(c => c.Hobbing && c.LineAB.Equals("CNC"));
                        //lấy ra các orders có TeethShape = Hob9CNC TeethShape
                        var ordersCNC_9 = _CNCOrdersHob.Where(c => c.TeethShape.Equals(Hob9CNC_TeethShape)).ToList();
                        //chia đôi số lượng items của CNCOrders
                        double chunkCNC = Math.Ceiling((_CNCOrdersHob.GroupBy(g => g.TeethShape).Count() * 1.0) / 2);

                        int hour = Convert.ToInt32(DateTime.Now.ToString("HH"));

                        //CNC orders with Hob9                          
                        if (ordersCNC_9.Count > 0)
                        {
                            //07-10-2020: test chuyển SortOrders_CNC sang function SetupSSD_CNC thời từ 6:00 AM 
                            if(hour >=6 && hour <24)
                            {
                                var sortedList = SetupSSD_CNC(ordersCNC_9, SSDType.N1,Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                                ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(sortedList);

                                //lấy ra phần còn lại kiểm tra nếu CNCHob9 có teethShape giống thì chia cho CNCHob10
                                var restOfListCNCHob9 = _CNCOrdersHob.Except(ListOrdersCNCHob9).ToList();
                                if (restOfListCNCHob9.Count > 0)
                                {
                                    //lấy hết phần còn lại của restOfListCNCHob9 đẩy về cho CNCHob10
                                    var lstSorted = SetupSSD_CNC(restOfListCNCHob9, SSDType.N1,Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(lstSorted);
                                }
                            }
                            else
                            {
                                //nếu có sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in                               
                                var sortedList = SortOrders_CNC(ordersCNC_9, Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                                ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(sortedList);

                                //lấy ra phần còn lại kiểm tra nếu CNCHob9 có teethShape giống thì chia cho CNCHob10
                                var restOfListCNCHob9 = _CNCOrdersHob.Except(ListOrdersCNCHob9).ToList();
                                if (restOfListCNCHob9.Count > 0)
                                {
                                    //lấy hết phần còn lại của restOfListCNCHob9 đẩy về cho CNCHob10
                                    var lstSorted = SortOrders_CNC(restOfListCNCHob9, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(lstSorted);
                                }
                            }
                                                   
                        }
                        else
                        {
                            
                            //nếu ko kiểm tra CNCHob10 có TeethShape và TeethQty giống hay ko
                            //lấy ra các orders có TeethShape = Hob10CNC TeethShape
                            var _CNCHob10orders = _CNCOrdersHob.Where(c => c.TeethShape.Equals(Hob10CNC_TeethShape)).ToList();
                            if (_CNCHob10orders.Count > 0)
                            {
                                //07-10-2020: test chuyển SortOrders_CNC sang function SetupSSD_CNC thời từ 6:00 AM 
                                if(hour >=6 && hour <24)
                                {
                                    //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                                    var _ordersCNCHob10 = SetupSSD_CNC(_CNCHob10orders, SSDType.N1,Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(_ordersCNCHob10);

                                    //phần còn lại sau khi chia cho CNCHob10 sẽ đẩy qua CNCHob9
                                    var theRestOfCNCHob10 = _CNCOrdersHob.Except(ListOrdersCNCHob10).ToList();
                                    if (theRestOfCNCHob10.Count > 0)
                                    {
                                        var _sorted_theRestOfCNCHob10 = SetupSSD_CNC(theRestOfCNCHob10, SSDType.N1,Hob10CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                                        ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(_sorted_theRestOfCNCHob10);
                                    }
                                }
                                else
                                {
                                    //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                                    var _ordersCNCHob10 = SortOrders_CNC(_CNCHob10orders, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(_ordersCNCHob10);

                                    //phần còn lại sau khi chia cho CNCHob10 sẽ đẩy qua CNCHob9
                                    var theRestOfCNCHob10 = _CNCOrdersHob.Except(ListOrdersCNCHob10).ToList();
                                    if (theRestOfCNCHob10.Count > 0)
                                    {
                                        var _sorted_theRestOfCNCHob10 = SortOrders_CNC(theRestOfCNCHob10, Hob10CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                                        ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(_sorted_theRestOfCNCHob10);
                                    }
                                }
                              
                            }
                            else
                            {
                                if(hour >= 6 && hour <24)
                                {
                                    //nếu ko chia đôi số lượng items của CNCOrders                              
                                    var tempLst = new List<OrderModel>();
                                    _CNCOrdersHob.GroupBy(g => g.TeethShape).Take((int)chunkCNC).Select(g => g.ToList()).ToList().ForEach(i => tempLst.AddRange(i));
                                    var sortedList = SetupSSD_CNC(tempLst,SSDType.N1);
                                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(sortedList);

                                    //lấy ra các orders có TeethShape = Hob10 CNC và không nằm trong ListOrdersCNCHob9
                                    _CNCOrdersHob.Where(c => ListOrdersCNCHob9.All(c2 => !c2.TeethShape.Equals(c.TeethShape))).ToList();
                                    var _ordersCNC_10 = _CNCOrdersHob.Except(ListOrdersCNCHob9).ToList();


                                    if (_ordersCNC_10.Count > 0)
                                    {
                                        //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                                        var sortedList2 = SetupSSD_CNC(_ordersCNC_10, SSDType.N1,Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(sortedList2);
                                    }

                                }
                                else
                                {
                                    //nếu ko chia đôi số lượng items của CNCOrders                              
                                    var tempLst = new List<OrderModel>();
                                    _CNCOrdersHob.GroupBy(g => g.TeethShape).Take((int)chunkCNC).Select(g => g.ToList()).ToList().ForEach(i => tempLst.AddRange(i));
                                    var sortedList = SortOrders_CNC(tempLst);
                                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(sortedList);

                                    //lấy ra các orders có TeethShape = Hob10 CNC và không nằm trong ListOrdersCNCHob9
                                    _CNCOrdersHob.Where(c => ListOrdersCNCHob9.All(c2 => !c2.TeethShape.Equals(c.TeethShape))).ToList();
                                    var _ordersCNC_10 = _CNCOrdersHob.Except(ListOrdersCNCHob9).ToList();


                                    if (_ordersCNC_10.Count > 0)
                                    {
                                        //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                                        var sortedList2 = SortOrders_CNC(_ordersCNC_10, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(sortedList2);
                                    }

                                }

                            }
                           
                        }
                       

                     
                        //CNC orders with No Hob                       
                        ListOrdersCNCNoHob = new ObservableCollection<OrderModel>(CNCOrders.Where(c => c.Hobbing == false && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => NoHobCNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));

                        BalancedOrdersCNC2();
                      

                    }
                 
                    //2019-11-27 Start add: tính total PO và total Orders cho CNC
                    Calc_PO_and_Orders_CNC();
                    //2019-11-27 End add

                    //2019-12-11 Start add: sau khi get Orders từ Manufa, enable tất cả các nút in
                    CNCHob9_IsGettingData = true;
                    CNCHob10_IsGettingData = true;
                    CNCNoHob_IsGettingData = true;
                    //2019-12-11 End add

                    //2019-11-14 Start add: sau khi lấy xong data cho CNC thì enable timerGetData để tự động lấy data mỗi giờ
                    timerGetData.Start();
                    //2019-11-14 End add

                    //sau khi lấy data manufa cho CNC
                    _mnfCNC_getData = false;
                    //all task finished
                    IsLoadingCNC = false;
                    IsGettingDataCNC = true;
                    timerCNCHob.Start();
                }
                catch (Exception ex)
                {
                    // Get stack trace for the exception with source file information
                    var st = new StackTrace(ex, true);
                    // Get the top stack frame
                    var frame = st.GetFrame(0);
                    // Get the line number from the stack frame
                    var line = frame.GetFileLineNumber();

                    MessageBox.Show("Máy chủ phản hồi quá lâu hoặc mất kết nối, vui lòng thử lại!\n"+line,"Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                }
              
            });
        }
        //2019-10-31 End add

        private bool CanGetOrdersBW(object obj)
        {
            if (BWHob11_IsPrinting || BWHob12_IsPrinting || BWNoHob_IsPrinting || BWCNCHob9_IsPrinting || BWCNCHob10_IsPrinting)
                return false;
            else return true;
        }


        //2019-11-1 Start add
        int _mnfCNCpage = 0;
        bool _mnfCNC_getData = false;
        //2019-11-1 End add
        private async void GetOrdersCNC(object parameters)
        {
            if (parameters != null)
            {

                await Task.Run(async () =>
                {
                    IsLoadingCNC = true;
                    timerCNCHob.Stop();
                    IsGettingDataCNC = false;

                    //2019-12-11 Start add: khi get Orders từ Manufa, disable tất cả các nút in
                    CNCHob9_IsGettingData = false;
                    CNCHob10_IsGettingData = false;
                    CNCNoHob_IsGettingData = false;
                    //2019-12-11 End add
                    var values = parameters as object[];

                    //WebBrowser wbCNCNoHob = values[1] as WebBrowser;
                    //WebBrowser wbCNCHob = values[0] as WebBrowser;

                    Awesomium.Windows.Controls.WebControl wbCNCHob = values[0] as Awesomium.Windows.Controls.WebControl;

                    //2019-11-14 Start add: gán button GetData cho việc lấy data tự động mỗi giờ
                    timerGetData.Stop();
                    btnGetDataCNC = values[3] as Button;
                    //2019-11-14 End add
                    await Task.Run(() =>
                    {

                        if (isFirstNavigationCNCNoHob)
                        {
                            //2019-11-11 Start lock: khóa tạm dùng chung 1 Awesomium thay cho Webbrowser
                            //navigate CNCNoHob page
                            //wbCNCNoHob.Dispatcher.InvokeAsync(() =>
                            //{
                            //    wbCNCNoHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                            //    wbCNCNoHob.LoadCompleted += new LoadCompletedEventHandler(wbCNCNoHob_LoadCompleted);
                            //});
                            //2019-11-11 End lock
                        }

                        if (isFirstNavigationCNCHob)
                        {
                            //navigate CNCHob page
                            wbCNCHob.Dispatcher.InvokeAsync(() =>
                            {
                                //wbCNCHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                                //wbCNCHob.LoadCompleted += new LoadCompletedEventHandler(wbCNCHob_LoadCompleted);
                                wbCNCHob.Source = new Uri("http://10.4.24.111:8080/manufa/Login.aspx");
                                wbCNCHob.LoadingFrameComplete += new Awesomium.Core.FrameEventHandler(wbCNCHob_LoadedFrameComplete);
                            });
                        }

                    });

                    await Task.Run(() =>
                    {
                        //2019-10-31 Start add: đổi cách lấy dữ liệu manufa
                        WebBrowser wbManufaDT = values[2] as WebBrowser;
                        //getting manufa data for CNC
                        _mnfCNC_getData = true;
                        wbManufaDT.Dispatcher.InvokeAsync(() =>
                        {
                            if (wbManufaDT != null)
                                wbManufaDT.LoadCompleted -= ManufaDT_LoadCompleted;

                            wbManufaDT.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");

                            wbManufaDT.LoadCompleted += new LoadCompletedEventHandler(ManufaDT_LoadCompleted);
                        });

                    });

                    //2019-11-1 Lock

                    ////get user login info in App.config
                    //string user = ConfigurationManager.AppSettings["user"];
                    //string pass = ConfigurationManager.AppSettings["pass"];

                    ////hide command prompt and browser when chromeDriver start
                    //var options = new ChromeOptions();
                    //options.AddArgument("--window-position=-32000,-32000");
                    //ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                    //service.HideCommandPromptWindow = true;
                    //ChromeDriver chromeDriver = new ChromeDriver(service, options);

                    ////Login manufa
                    //await Task.Run(() =>
                    //{
                    //    chromeDriver.Url = "http://10.4.24.111:8080/manufa/Login.aspx";
                    //    chromeDriver.Navigate();
                    //    Task.Delay(5000).Wait();

                    //    if (isFirstNavigationCNCNoHob)
                    //    {
                    //        //navigate CNCNoHob page
                    //        wbCNCNoHob.Dispatcher.InvokeAsync(() =>
                    //        {
                    //            wbCNCNoHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                    //            wbCNCNoHob.LoadCompleted += new LoadCompletedEventHandler(wbCNCNoHob_LoadCompleted);
                    //        });
                    //    }

                    //    if (isFirstNavigationCNCHob)
                    //    {
                    //        //navigate CNCHob page
                    //        wbCNCHob.Dispatcher.InvokeAsync(() =>
                    //        {
                    //            wbCNCHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                    //            wbCNCHob.LoadCompleted += new LoadCompletedEventHandler(wbCNCHob_LoadCompleted);
                    //        });
                    //    }

                    //});
                    //PreparingCNCData(chromeDriver);

                    //2019-11-1 End lock
                });


            }
        }

        private bool CanGetOrdersCNC(object obj)
        {
            if (CNCHob9_IsPrinting || CNCHob10_IsPrinting || CNCNoHob_IsPrinting)
                return false;
            else return true;
        }

        //KEY-CNC-1
        private async void PreparingCNCData(ChromeDriver chromeDriver)
        {
            await Task.Run(async () =>
            {
                //delete existed Excel
                //download folder path
                string downloadFolderPath = new KnownFolder(KnownFolderType.Downloads).Path;
                string[] files = Directory.GetFiles(downloadFolderPath);
                try
                {
                    foreach (var file in files)
                    {
                        if (file.Contains("OrdersList"))
                        {
                            File.Delete(file);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    
                    MessageBox.Show("OrdersList excel file is being used, please close it and and try again!", "Message", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }


                //login info
                var userInput = chromeDriver.FindElementById("txtUserId");
                var passInput = chromeDriver.FindElementById("txtPassword");
                var btnLogin = chromeDriver.FindElementById("cmdLogin");
                //login input
                //get user login info in App.config
                string user = ConfigurationManager.AppSettings["user"];
                string pass = ConfigurationManager.AppSettings["pass"];
                userInput.SendKeys(user);
                passInput.SendKeys(pass);
                btnLogin.Click();

                //wait until page is load
                chromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);

                //navigate manufacturing
                var manufacturing_Instructions_Print = chromeDriver.FindElementByXPath("//*[@id=\"MenuDetail_0\"]/tbody/tr[1]/td");
                manufacturing_Instructions_Print.Click();

                //wait until page is load
                chromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);

                //navigate orders list
                var forms = chromeDriver.FindElementById("ContentPlaceHolder1_drpSelectForm");
                SelectElement selectForm = new SelectElement(forms);
                selectForm.SelectByValue("Orders List");

                //wait until page is load
                chromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);

                //input search filter
                var statuses = chromeDriver.FindElementById("ContentPlaceHolder1_drpStatus");
                SelectElement selectStatus = new SelectElement(statuses);
                selectStatus.SelectByValue("Authorization already (Not Print)");

                //wait until page is load
                chromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);

                //input search filter
                var line = chromeDriver.FindElementById("ContentPlaceHolder1_txtLine");
                line.SendKeys("LE");
                //onclick search
                var btnSearch = chromeDriver.FindElementById("ContentPlaceHolder1_cmbSelect");
                btnSearch.Click();

                //wait until page is load
                chromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);

                //download file excel
                var btnExcelOutput = chromeDriver.FindElementById("ContentPlaceHolder1_cmbCsvOutput");
                btnExcelOutput.Click();

                Task.Delay(5000).Wait();

                //dispose chromeDriver 
                chromeDriver.Close();
                chromeDriver.Quit();

                //read file Excel
                string excelFile = downloadFolderPath + "\\OrdersList.xlsx";
                List<OrderModel> lstManufaInfo = GetCNCOrdersFromExcel(excelFile).ToList();
                if (lstManufaInfo.Count > 0)
                {
                    //get InnerCode from lstManufaInfo 
                    var innerCodes = string.Format("'{0}'", string.Join("','", lstManufaInfo.Select(c => c.Inner_Code)));
                    List<OrderModel> lstPulleyMaster = OrderBLL.Instance.GetPulleyMasterByProduct(innerCodes);
                    //get LineAB
                    //2019-10-05 Start add: get machine types to identify which PO belongs to BW or CNC
                    List<MachineTypeModel> lstMachineTypes = await MachineTypeBLL.Instance.GetMachineTypes();
                    //2019-10-05 End add
                    if (lstPulleyMaster.Count > 0)
                    {


                        IEnumerable<OrderModel> result = (from r1 in lstManufaInfo
                                                          join r2 in lstPulleyMaster on r1.Inner_Code equals r2.Inner_Code
                                                          join r3 in lstMachineTypes on r1.Material_code1 equals r3.MaterialCode
                                                          select new OrderModel()
                                                          {
                                                              //No = r1.No,
                                                              Sales_Order_No = r1.Sales_Order_No,
                                                              Pier_Instruction_No = r1.Pier_Instruction_No,
                                                              Manufa_Instruction_No = r1.Manufa_Instruction_No,
                                                              Global_Code = r1.Global_Code,
                                                              //Customers = r1.Customers,
                                                              Item_Name = r1.Item_Name,
                                                              MC = r1.MC,
                                                              Received_Date = r1.Received_Date,
                                                              Factory_Ship_Date = r1.Factory_Ship_Date,
                                                              Number_of_Orders = r1.Number_of_Orders,
                                                              Number_of_Available_Instructions = r1.Number_of_Available_Instructions,
                                                              //Number_of_Repairs = r1.Number_of_Repairs,
                                                              //Number_of_Instructions = r1.Number_of_Instructions,
                                                              Line = r1.Line,
                                                              //PayWard = r1.PayWard,
                                                              //Major = r1.Major,
                                                              //Special_Orders = r1.Special_Orders,
                                                              //Method = r1.Method,
                                                              //Destination = r1.Destination,
                                                              //Instructions_Print_date = r1.Instructions_Print_date,
                                                              //Latest_progress = r1.Latest_progress,
                                                              //Tack_Label_Output_Date = r1.Tack_Label_Output_Date,
                                                              //Completion_Instruction_Date = r1.Completion_Instruction_Date,
                                                              //Re_print_Count = r1.Re_print_Count,
                                                              //Latest_issue_time = r1.Latest_issue_time,
                                                              Material_code1 = r1.Material_code1,
                                                              Material_text1 = r1.Material_text1.Substring(r1.Material_text1.LastIndexOf("/") + 1, 3),
                                                              Amount_used1 = r1.Amount_used1,
                                                              Unit1 = r1.Unit1,
                                                              //Material_code2 = r1.Material_code2,
                                                              //Material_text2 = r1.Material_text2,
                                                              //Amount_used2 = r1.Amount_used2,
                                                              //Unit2 = r1.Unit2,
                                                              //Material_code3 = r1.Material_code3,
                                                              //Material_text3 = r1.Material_text3,
                                                              //Amount_used3 = r1.Amount_used3,
                                                              //Unit3 = r1.Unit3,
                                                              //Material_code4 = r1.Material_code4,
                                                              //Material_text4 = r1.Material_text4,
                                                              //Amount_used4 = r1.Amount_used4,
                                                              //Unit4 = r1.Unit4,
                                                              //Material_code5 = r1.Material_code5,
                                                              //Material_text5 = r1.Material_text5,
                                                              //Amount_used5 = r1.Amount_used5,
                                                              //Unit5 = r1.Unit5,
                                                              //Material_code6 = r1.Material_code6,
                                                              //Material_text6 = r1.Material_text6,
                                                              //Amount_used6 = r1.Amount_used6,
                                                              //Unit6 = r1.Unit6,
                                                              //Material_code7 = r1.Material_code7,
                                                              //Material_text7 = r1.Material_text7,
                                                              //Amount_used7 = r1.Amount_used7,
                                                              //Unit7 = r1.Unit7,
                                                              //Material_code8 = r1.Material_code8,
                                                              //Material_text8 = r1.Material_text8,
                                                              //Amount_used8 = r1.Amount_used8,
                                                              //Unit8 = r1.Unit8,
                                                              Inner_Code = r1.Inner_Code,
                                                              //Classify_Code = r1.Classify_Code,
                                                              Hobbing = r2.Hobbing,
                                                              TeethQuantity = r2.TeethQuantity,
                                                              TeethShape = r2.TeethShape,
                                                              LineAB = r3.LineAB,
                                                              LineName = r3.LineName
                                                          }).ToList();

                        CNCOrders = new ObservableCollection<OrderModel>(result);
                        //lấy ra các orders có TeethShape = Hob9CNC TeethShape
                        var ordersCNC_9 = CNCOrders.Where(c => c.TeethShape.Equals(Hob9CNC_TeethShape)).ToList();
                        //chia đôi số lượng items của CNCOrders
                        double chunkCNC = Math.Ceiling((CNCOrders.GroupBy(g => g.TeethShape).Count() * 1.0) / 2);

                        //CNC orders with Hob9   
                        //2019-10-30 Start lock
                        //ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(CNCOrders.Where(c => c.Hobbing && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => Hob9CNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity == Hob9CNC_TeethQty).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape));
                        //2019-10-30 End lock

                        //2019-10-30 Start add
                        if (ordersCNC_9.Count > 0)
                        {
                            //nếu có sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in                               
                            var sortedList = SortOrders(ordersCNC_9, Hob9CNC_TeethShape, Hob9CNC_TeethQty);
                            ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(sortedList);
                        }
                        else
                        {
                            //nếu ko chia đôi số lượng items của CNCOrders   
                            var tempList = new List<OrderModel>();
                            CNCOrders.GroupBy(g => g.TeethShape).Take((int)chunkCNC).Select(g => g.ToList()).ToList().ForEach(i => tempList.AddRange(i));
                            var sortedList = SortOrders(tempList);
                            ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(sortedList);
                        }
                        //2019-10-30 End add

                        //CNC orders with Hob10        
                        //2019-10-30 Start lock
                        //ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(CNCOrders.Where(c=> ListOrdersCNCHob9.All(c2 => !c2.TeethShape.Equals(c.TeethShape))).Where(c => c.Hobbing && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => Hob10CNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                        //2019-10-30 End lock

                        //2019-10-29 End add

                        //lấy ra các orders có TeethShape = Hob10 CNC và không nằm trong ListOrdersCNCHob10                        
                        var ordersCNC_10 = CNCOrders.Except(ListOrdersCNCHob9).ToList();

                        //2019-10-29 Start add
                        if (ordersCNC_10.Count > 0)
                        {
                            //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                            var sortedList = SortOrders(ordersCNC_10, Hob10CNC_TeethShape, Hob10CNC_TeethQty);
                            ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(sortedList);
                        }
                        //2019-10-29 End add  


                        //CNC orders with No Hob                       
                        ListOrdersCNCNoHob = new ObservableCollection<OrderModel>(CNCOrders.Where(c => c.Hobbing == false && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => NoHobCNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));

                        /*
                        //2019-10-29 Start add                        
                        if (ordersBW_11.Count > 0)
                        {
                            //nếu có sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in                               
                            var sortedList = SortOrders(ordersBW_11, Hob11BW_TeethShape, Hob11BW_TeethQty);
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedList);
                        }
                        else
                        {
                            //nếu ko chia đôi số lượng items của CNCOrders                              
                            var sortedList = SortOrders(_BWOrders.Take((int)chunkBW).ToList());
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedList);
                        }

                        //2019-10-29 End add

                        //lấy ra các orders có TeethShape = Hob12 BW và không nằm trong ListOrdersBWHob11
                        //_BWOrders.Where(c => ListOrdersBWHob11.All(c2 => !c2.TeethShape.Equals(c.TeethShape))).ToList();
                        var ordersBW_12 = _BWOrders.Except(ListOrdersBWHob11).ToList();
                        
                        //2019-10-29 Start add
                        if(ordersBW_12.Count>0)
                        {
                            //sắp xếp TeethQuantity gần nhất với TeethQuantity trước đó đã in
                            var sortedList = SortOrders(ordersBW_12, Hob12BW_TeethShape, Hob12BW_TeethQty);
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(sortedList);
                        }
                        //2019-10-29 End add    
                     */
                    }

                }
                //all task finished
                IsLoadingCNC = false;
                IsGettingDataCNC = true;

            });
        }
        //KEY-CNC-2
        private void wbCNCHob_LoadedFrameComplete(object sender, Awesomium.Core.FrameEventArgs e)
        {
            isFirstNavigationCNCHob = false;
            Awesomium.Windows.Controls.WebControl wb = sender as Awesomium.Windows.Controls.WebControl;
            dynamic document = (Awesomium.Core.JSObject)wb.ExecuteJavascriptWithResult("document");
            using (document)
            {
                if (isFirstLoadCNCHob)
                {

                    //get user login info in App.config
                    string user = ConfigurationManager.AppSettings["user"];
                    string pass = ConfigurationManager.AppSettings["pass"];

                    //login
                    dynamic userInput = document.getElementById("txtUserId");
                    dynamic passInput = document.getElementById("txtPassword");
                    dynamic btnLogin = document.getElementById("cmdLogin");
                    userInput.setAttribute("value", user);
                    passInput.setAttribute("value", pass);
                    btnLogin.click();
                    isFirstLoadCNCHob = false;
                    //điều hướng trang về 0 đề vào lại trang in
                    pageNumberCNCHob = 0;
                }
                else
                {

                    if (pageNumberCNCHob == 0)
                    {
                        //nếu trang tải xong
                        if (wb.IsLoaded)
                        {
                            //check session time out
                            if (CheckSessionExpiredAwesomiumBrowser(wb))
                            {
                                //navigate to login page
                                isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                                document.getElementById("btnOK").click();
                            }
                            else
                            {
                                //still on the session
                                var navigate = wb.ExecuteJavascriptWithResult("window.location='/manufa/PrintViewInstruction.aspx';");
                                //chuyển trang
                                pageNumberCNCHob = 1;
                            }
                        }

                    }
                    else if (pageNumberCNCHob == 1)
                    {
                        //check session time out
                        if (CheckSessionExpiredAwesomiumBrowser(wb))
                        {
                            //2019-10-11 Start add: if CNCHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                            _CNCHob_IsOutOfSession = true;
                            //2019-10-11 End add

                            //navigate to login page
                            isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {
                            //2019-10-11 Start add
                            if (_CNCHob_IsOutOfSession)
                            {
                                //_CNCHob_IsOutOfSession = false;//reset to default value
                                if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                                {
                                    SelectedOrderCNCHob9 = ListOrdersCNCHob9[0];
                                }
                                else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                                {
                                    SelectedOrderCNCHob10 = ListOrdersCNCHob10[0];
                                }
                               
                                //onclick btnClear to navigate to page 3
                                pageNumberCNCHob = 2;
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                            }
                            //2019-10-11 End add
                            else
                            {
                                var printMode = wb.ExecuteJavascriptWithResult(@"                                                                                                          
                                        //set selected status option for the order that not printed yet
                                        var printMode = document.getElementById('ContentPlaceHolder1_rdoSelectOne');
                                        printMode.checked = true;");

                                if (printMode)
                                    pageNumberCNCHob = 2;

                            }
                        }

                    }
                    else if (pageNumberCNCHob == 2)
                    {
                        //2019-11-14 Start add: reset lại trạng thái session và trạng thái nút PRINT sau khi re-login, thông báo in lại
                        if (_CNCHob_IsOutOfSession)
                        {
                            _CNCHob_IsOutOfSession = false;
                            IsGettingDataCNC = true;
                            if (CNCHob9_IsPrinting)
                            {
                                CNCHob9_IsGettingData = true;
                                CNCHob9_IsPrinting = false;
                                CNC_Hob9Status = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                            }
                            else if (CNCHob10_IsPrinting)
                            {
                                CNCHob10_IsGettingData = true;
                                CNCHob10_IsPrinting = false;
                                CNC_Hob10Status = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                            }
                            else
                            {
                                CNCNoHob_IsGettingData = true;
                                CNCNoHob_IsPrinting = false;
                                CNC_NoHobStatus = "Hệ thống đăng nhập lại thành công, bạn vui lòng in lại!";
                            }
                        }
                        //2019-11-14 End add
                        //2019-11-05 Start lock: bỏ bớt bước này, đưa PO vào Manufa search ngay khi nút "Print" được ấn
                        //Disable btnGetData, enable BWHobLoading while the printer is running                    
                        //if (BWHob11_IsPrinting) //nếu BWHob11 đang in
                        //{
                        //    BWHob11_IsGettingData = false;
                        //}
                        //else if (BWHob12_IsPrinting) //nếu BWHob12 đang in
                        //{
                        //    BWHob12_IsGettingData = false;
                        //}
                        //else if (BWCNCHob9_IsPrinting) //nếu BWCNCHob9 đang in
                        //{
                        //    BWCNCHob9_IsGettingData = false;
                        //}
                        //else if (BWCNCHob10_IsPrinting)
                        //{
                        //    BWCNCHob10_IsGettingData = false;
                        //}

                        //if (wb.IsLoaded)
                        //{

                        //    if (CheckSessionExpired(wb))
                        //    {
                        //        //2019-10-11 Start add: if BWHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        //        _BWHob_IsOutOfSession = true;
                        //        //2019-10-11 End add

                        //        //navigate to login page
                        //        isFirstLoadBWHob = true; //true: login page <> false: other pages     
                        //        document.getElementById("btnOK").click();

                        //    }
                        //    else
                        //    {
                        //        //điền PO (manufa NO) và "Search" đơn để in
                        //        if (BWHob11_IsPrinting) //nếu BWHob11 đang in
                        //        {
                        //            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob11.Manufa_Instruction_No);
                        //            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                        //            pageNumberBWHob = 3;
                        //        }
                        //        else if (BWHob12_IsPrinting) //nếu BWHob12 đang in
                        //        {
                        //            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob12.Manufa_Instruction_No);
                        //            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                        //            pageNumberBWHob = 3;
                        //        }
                        //        else if (BWCNCHob9_IsPrinting) //nếu BWCNCHob9 đang in
                        //        {
                        //            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWCNCHob9.Manufa_Instruction_No);
                        //            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                        //            pageNumberBWHob = 3;
                        //        }
                        //        else if (BWCNCHob10_IsPrinting)
                        //        {
                        //            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWCNCHob10.Manufa_Instruction_No);
                        //            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                        //            pageNumberBWHob = 3;
                        //        }

                        //    }
                        //}
                        //2019-11-05 End lock
                    }
                    else if (pageNumberCNCHob == 3)
                    {
                        if (CheckSessionExpiredAwesomiumBrowser(wb))
                        {
                            //2019-10-11 Start add: if CNCHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                            _CNCHob_IsOutOfSession = true;
                            //2019-10-11 End add

                            //thông báo hết session
                            if (CNCHob9_IsPrinting)
                            {
                                CNC_Hob9Status = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                            }
                            else if (CNCHob10_IsPrinting)
                            {
                                CNC_Hob10Status = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                            }
                            else
                            {
                                CNC_NoHobStatus = "Session hết hạn, hệ thống đang đăng nhập lại, bạn vui lòng chờ...";
                            }

                            //navigate to login page
                            isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {
                            
                             var print = wb.ExecuteJavascriptWithResult(@"          
                                                                                    
                                        var orderToPrint = document.getElementById('ContentPlaceHolder1_grdPrintList_chkOutput_0');
                                        var btnPrint = document.getElementById('ContentPlaceHolder1_cmdPrint');
                                            
                                        if(orderToPrint != null)
                                        {   
                                            //check vào PO cần in
                                            orderToPrint.checked = 'checked';   
                                            //chọn máy in, 16 = pulley máy 3
                                            var printer = document.getElementById('ContentPlaceHolder1_drpPrinter');
                                            //2019-12-05 Manufa thay đổi vị trí máy in
                                            //printer.selectedIndex = 16
                                            printer.value = 122; 
                                            //pulley 03 value = 122;
                                            //2019-12-05
                                            //in 
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },500);
                                            

                                        }");
                          
                                pageNumberCNCHob = 4; //chuyển trang sau khi in, lấy status
                           
                        }

                    }
                    else if (pageNumberCNCHob == 4)
                    {
                        //2019-11-06 Start add: kiểm tra PO được in hay ko
                        if (wb.IsLoaded)
                        {
                            if (CheckSessionExpiredAwesomiumBrowser(wb))
                            {
                                //navigate to login page
                                isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                                document.getElementById("btnOK").click();

                            }
                            else
                            {
                                //lấy status sau khi in
                                string status = document.getElementById("ContentPlaceHolder1_AlertArea").innerHTML;
                                //nếu in thành công
                                if (status.ToLower().Contains("completed"))
                                {
                                    if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                                    {
                                        Hob9CNC_TeethShape = SelectedOrderCNCHob9.TeethShape;
                                        Hob9CNC_TeethQty = SelectedOrderCNCHob9.TeethQuantity;
                                        Hob9CNC_Diameter = SelectedOrderCNCHob9.Diameter_d;
                                        Hob9CNC_DMat = SelectedOrderCNCHob9.Material_text1;
                                        Hob9CNC_GlobalCode = SelectedOrderCNCHob9.Global_Code;
                                        //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                        var poToDelete = ListOrdersCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob9.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersCNCHob9.Remove(poToDelete);
                                            //2019-12-02 Start change: đổi cách lưu thông tin PO cho việc đồng bộ CNCHob9 và BWCNCHob9
                                            //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob9.json");
                                            string query = string.Format("UPDATE CNCHob9 SET SoPO='{0}', TeethShape='{1}', TeethQty={2}, Diameter_d={3}, DMat='{4}', GlobalCode = '{5}'", poToDelete.Manufa_Instruction_No, poToDelete.TeethShape, poToDelete.TeethQuantity, poToDelete.Diameter_d, poToDelete.Material_text1, poToDelete.Global_Code);
                                            try
                                            {
                                                MySQLProvider.Instance.ExecuteNonQuery(query);
                                                CNC_Hob9Status = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                                Task.Delay(1000).Wait();
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("In PO thất bại, vui lòng thử lại!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                            }
                                            //2019-12-02 End change
                                           
                                            
                                            //2019-10-15 Start add: enable timer again
                                            timerCNCHob.Start();
                                            //2019-10-15 End add
                                        }
                                        //2019-11-04 Start lock
                                        //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                        //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                        //btnClear.click();
                                        //2019-10-04 End add
                                        //2019-11-04 End lock

                                        //2019-11-05 Start lock: thay đổi trạng thái nút PRINT sau in
                                        //2019-10-07 Start add
                                        //Enable btnGetData, disable CNCHobLoading while the printer is running
                                        IsGettingDataCNC = true;
                                        CNCHob9_IsGettingData = true;
                                        CNCHob9_IsPrinting = false;
                                        //2019-10-07 End add
                                        //2019-11-05 End lock

                                        //2019-11-08 Start lock: khóa tạm thời chức năng điền tiếp PO sau khi in
                                        //2019-11-04 Start add
                                        //if (ListOrdersBWHob11 != null && ListOrdersBWHob11.Count > 0)
                                        //{
                                        //    SelectedOrderBWHob11 = ListOrdersBWHob11.FirstOrDefault();
                                        //    //sau khi in thành công search tiếp PO kế tiếp và in (người dùng ko cần dùng tay click nút PRINT)
                                        //    document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob11.Manufa_Instruction_No);
                                        //    document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                                        //    //chuyển hướng trang 5 để PO sau khi search để PO sẵn sàng in
                                        //    pageNumberBWHob11 = 5;
                                        //}
                                        //else
                                        //{
                                        //    //nếu hết đơn thì trả về trạng thái default cho btnPRINT
                                        //    IsGettingDataBW = true;
                                        //    BWHob11_IsGettingData = true;
                                        //    BWHob11_IsPrinting = false;

                                        //    //chuyển hướng trang 
                                        //    pageNumberBWHob11 = -1;
                                        //}                                    
                                        //2019-11-04 End add
                                        //2019-11-08 End lock
                                    }
                                    else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                                    {
                                        Hob10CNC_TeethShape = SelectedOrderCNCHob10.TeethShape;
                                        Hob10CNC_TeethQty = SelectedOrderCNCHob10.TeethQuantity;
                                        Hob10CNC_Diameter = SelectedOrderCNCHob10.Diameter_d;
                                        Hob10CNC_DMat = SelectedOrderCNCHob10.Material_text1;
                                        Hob10CNC_GlobalCode = SelectedOrderCNCHob10.Global_Code;
                                        //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                        var poToDelete = ListOrdersCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob10.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersCNCHob10.Remove(poToDelete);
                                            //2019-12-02 Start change: đổi cách lưu thông tin PO cho việc đồng bộ giữa CNCHob10 và BWCNCHob10
                                            //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob10.json");
                                            string query = string.Format("UPDATE CNCHob10 SET SoPO='{0}', TeethShape='{1}', TeethQty={2}, Diameter_d={3}, DMat = '{4}', GlobalCode = '{5}'", poToDelete.Manufa_Instruction_No, poToDelete.TeethShape, poToDelete.TeethQuantity, poToDelete.Diameter_d, poToDelete.Material_text1, poToDelete.Global_Code);
                                            try
                                            {
                                                MySQLProvider.Instance.ExecuteNonQuery(query);
                                                CNC_Hob10Status = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                                Task.Delay(1000).Wait();
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("In PO thất bại, vui lòng thử lại!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                            }
                                            //2019-12-02 End change
                                           
                                            
                                            //2019-10-15 Start add: enable timer again
                                            timerCNCHob.Start();
                                            //2019-10-15 End add
                                        }
                                        //2019-11-04 Start lock
                                        //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                        //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                        //btnClear.click();
                                        //2019-10-04 End add
                                        //2019-11-04 End lock

                                        //2019-11-05 Start lock
                                        //2019-10-07 Start add
                                        //Enable btnGetData, disable BWHobLoading while the printer is running
                                        IsGettingDataCNC = true;
                                        CNCHob10_IsGettingData = true;
                                        CNCHob10_IsPrinting = false;
                                        //2019-10-07 End add
                                        //2019-11-05 End lock

                                        //2019-11-08 Start lock: khóa tạm thời chức năng điền tiếp PO sau khi in
                                        //2019-11-04 Start add
                                        //if (ListOrdersBWHob12 != null && ListOrdersBWHob12.Count > 0)
                                        //{
                                        //    SelectedOrderBWHob12 = ListOrdersBWHob12.FirstOrDefault();
                                        //    //sau khi in thành công search tiếp PO kế tiếp và in (người dùng ko cần dùng tay click nút PRINT)
                                        //    document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob12.Manufa_Instruction_No);
                                        //    document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                                        //    //chuyển hướng trang sau khi search
                                        //    pageNumberBWHob11 = 3;
                                        //}
                                        //else
                                        //{
                                        //    //nếu hết đơn thì trả về trạng thái default cho btnPRINT
                                        //    IsGettingDataBW = true;
                                        //    BWHob12_IsGettingData = true;
                                        //    BWHob12_IsPrinting = false;
                                        //}

                                        //2019-11-04 End add
                                        //2019-11-08 End lock
                                    }
                                    else if (CNCNoHob_IsPrinting) //nếu CNCNoHob đang in
                                    {
                                        NoHobCNC_TeethShape = SelectedOrderCNCNoHob.TeethShape;
                                        NoHobCNC_TeethQty = SelectedOrderCNCNoHob.TeethQuantity;
                                        NoHobCNC_Diameter = SelectedOrderCNCNoHob.Diameter_d;
                                        //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                        var poToDelete = ListOrdersCNCNoHob.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersCNCNoHob.Remove(poToDelete);

                                            CNC_NoHobStatus = "PO: " + poToDelete.Manufa_Instruction_No + " has been completed";
                                            Task.Delay(1000).Wait();
                                            //2019-10-15 Start add: enable timer again
                                            timerCNCHob.Start();
                                            //2019-10-15 End add
                                        }
                                        //2019-11-04 Start lock
                                        //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                        //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                        //btnClear.click();
                                        //2019-10-04 End add
                                        //2019-11-04 End lock

                                        //2019-10-07 Start add
                                        //Enable btnGetData, disable BWHobLoading while the printer is running
                                        IsGettingDataBW = true;
                                        CNCNoHob_IsGettingData = true;
                                        CNCNoHob_IsPrinting = false;
                                        //2019-10-07 End add

                                        //2019-11-04 Start add
                                        //pageNumberBWHob11 = 5;//in tiếp
                                        //2019-11-04 End add
                                    }

                                }
                                else if (status.ToLower().Contains("found"))
                                {
                                    //Ko tìm thấy PO
                                    //PO chua duoc chon
                                    if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                                    {
                                        //xóa PO đã in thành công khỏi list
                                        var poToDelete = ListOrdersCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob9.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersCNCHob9.Remove(poToDelete);
                                            //2019-10-11 Start add: print message: Order has been printed
                                            CNC_Hob9Status = "PO: " + poToDelete + " has been printed!";
                                        }
                                        //2019-10-07 End add
                                        CNCHob9_IsGettingData = true;
                                        CNCHob9_IsPrinting = false;
                                    }
                                    else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                                    {
                                        //xóa PO đã in thành công khỏi list
                                        var poToDelete = ListOrdersCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob10.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersCNCHob10.Remove(poToDelete);
                                            //2019-10-11 Start add: print message: Order has been printed
                                            CNC_Hob10Status = "PO: " + poToDelete + " has been printed!";
                                        }
                                        //2019-10-07 End add
                                        CNCHob10_IsGettingData = true;
                                        CNCHob10_IsPrinting = false;
                                    }
                                    else if (CNCNoHob_IsPrinting) //nếu CNCNoHob đang in
                                    {
                                        //xóa PO đã in thành công khỏi list
                                        var poToDelete = ListOrdersCNCNoHob.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                        if (poToDelete != null)
                                        {
                                            ListOrdersCNCNoHob.Remove(poToDelete);
                                            //2019-10-11 Start add: print message: Order has been printed
                                            CNC_NoHobStatus = "PO: " + poToDelete + " has been printed!";
                                        }
                                        //2019-10-07 End add
                                        CNCNoHob_IsGettingData = true;
                                        CNCNoHob_IsPrinting = false;
                                    }


                                    //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    btnClear.click();
                                    //2019-10-04 End add


                                    //Enable btnGetData
                                    IsGettingDataCNC = true;

                                    //2019-10-07 End add
                                    //chuyển hướng trang 
                                    pageNumberCNCHob = -1;
                                }
                                else
                                {
                                    //máy in hết giấy, offline, etc
                                    //quay về trang in
                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable CNCHobLoading if getting error
                                    IsGettingDataCNC = true;
                                    if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                                    {
                                        //2019-10-11 Start add: print message: Order has been printed
                                        CNC_Hob9Status = "Printer offline or paper ran out, please check and try again!";
                                        //2019-10-07 End add
                                        CNCHob9_IsGettingData = true;
                                        CNCHob9_IsPrinting = false;
                                    }
                                    else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                                    {
                                        //2019-10-11 Start add: print message: Order has been printed
                                        CNC_Hob10Status = "Printer offline or paper ran out, please check and try again!";
                                        //2019-10-07 End add
                                        CNCHob10_IsGettingData = true;
                                        CNCHob10_IsPrinting = false;
                                    }
                                    else if (CNCNoHob_IsPrinting) //nếu CNCNoHob đang in
                                    {
                                        //2019-10-11 Start add: print message: Order has been printed
                                        CNC_NoHobStatus = "Printer offline or paper ran out, please check and try again!";
                                        //2019-10-07 End add
                                        CNCNoHob_IsGettingData = true;
                                        CNCNoHob_IsPrinting = false;
                                    }
                                    //2019-10-07 End add

                                    //chuyển hướng trang 
                                    pageNumberCNCHob = -1;
                                }

                            }

                            //2019-11-27 Start add: tính total PO và total Orders cho CNC
                            Calc_PO_and_Orders_CNC();
                            //2019-11-27 End add
                        }
                        //2019-11-06 End add

                    }
                    else if (pageNumberBWHob11 == 5)
                    {
                      
                    }
                    

                }

            }

        }

        private void wbCNCHob_LoadCompleted(object sender, NavigationEventArgs e)
        {
            isFirstNavigationCNCHob = false;
            WebBrowser wb = sender as WebBrowser;
            mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;
            if (isFirstLoadCNCHob)
            {

                //get user login info in App.config
                string user = ConfigurationManager.AppSettings["user"];
                string pass = ConfigurationManager.AppSettings["pass"];

                //login
                mshtml.IHTMLElement userInput = document.getElementById("txtUserId");
                mshtml.IHTMLElement passInput = document.getElementById("txtPassword");
                mshtml.IHTMLElement btnLogin = document.getElementById("cmdLogin");
                userInput.setAttribute("value", user);
                passInput.setAttribute("value", pass);
                btnLogin.click();
                isFirstLoadCNCHob = false;
            }
            else
            {

                if (pageNumberCNCHob == 0)
                {
                    //nếu trang tải xong
                    if (wb.IsLoaded)
                    {
                        //check session time out
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();
                        }
                        else
                        {
                            //still on the session
                            var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                            var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                            script.text = @" 
                                   window.onload = function(){
                                        //navigate 'Manufacturing Instructions Print' menu
                                        window.location = '/manufa/PrintViewInstruction.aspx';
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);
                            HideScriptErrors(wb, true);
                            pageNumberCNCHob = 1;
                        }
                    }

                }
                else if (pageNumberCNCHob == 1)
                {
                    //check session time out
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if CNCHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _CNCHob_IsOutOfSession = true;
                        //2019-10-11 End add

                        //navigate to login page
                        isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();

                    }
                    else
                    {
                        //2019-10-11 Start add
                        if (_CNCHob_IsOutOfSession)
                        {
                            _CNCHob_IsOutOfSession = false;//reset to default value
                            if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                            {
                                SelectedOrderCNCHob9 = ListOrdersCNCHob9[0];
                            }
                            else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                            {
                                SelectedOrderCNCHob10 = ListOrdersCNCHob10[0];
                            }


                            //onclick btnClear to navigate to page 2
                            pageNumberCNCHob = 2;
                            var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                            btnClear.click();
                        }
                        //2019-10-11 End add
                        else
                        {
                            var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                            var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                            script.text = @" 
                                   window.onload = function(){                                                                              
                                        //set selected status option for the order that not printed yet
                                        var printMode = document.getElementById('ContentPlaceHolder1_rdoSelectOne');
                                        printMode.checked = true;        
                                          
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);
                            pageNumberCNCHob = 2;

                        }

                    }

                }
                else if (pageNumberCNCHob == 2)
                {

                    //Disable btnGetData, enable CNCHobLoading while the printer is running   
                    if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                    {
                        CNCHob9_IsGettingData = false;
                    }
                    else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                    {
                        CNCHob10_IsGettingData = false;
                    }


                    if (wb.IsLoaded)
                    {

                        if (CheckSessionExpired(wb))
                        {
                            //2019-10-11 Start add: if CNCHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                            _CNCHob_IsOutOfSession = true;
                            //2019-10-11 End add

                            //navigate to login page
                            isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {


                            //điền PO (manufa NO) và "Search" đơn để in
                            if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                            {

                                document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCHob9.Manufa_Instruction_No);
                                document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                                pageNumberCNCHob = 3;
                            }
                            else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                            {
                                document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCHob10.Manufa_Instruction_No);
                                document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                                pageNumberCNCHob = 3;
                            }

                        }
                    }
                }
                else if (pageNumberCNCHob == 3)
                {
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if CNCHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _CNCHob_IsOutOfSession = true;
                        //2019-10-11 End add

                        //navigate to login page
                        isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();

                    }
                    else
                    {

                        var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                        var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                        script.text = @" 
                                   window.onload = function(){                
                                                                                    
                                        var orderToPrint = document.getElementById('ContentPlaceHolder1_grdPrintList_chkOutput_0');
                                        var btnPrint = document.getElementById('ContentPlaceHolder1_cmdPrint');
                                            
                                        if(orderToPrint != null)
                                        {   
                                            //check vào PO cần in
                                            orderToPrint.checked = 'checked';   
                                            //chọn máy in, 4 = pulley
                                            var printer = document.getElementById('ContentPlaceHolder1_drpPrinter');
                                            printer.selectedIndex = 4
                                            //in 
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },2000);
                                            
                                            
                                        }
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        pageNumberCNCHob = 4; //chuyển trang sau khi in, lấy status
                    }


                }
                else if (pageNumberCNCHob == 4)
                {
                    if (wb.IsLoaded)
                    {
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadCNCHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {
                            //lấy status sau khi in
                            string status = document.getElementById("ContentPlaceHolder1_AlertArea").innerHTML;
                            //nếu in thành công
                            if (status.ToLower().Contains("completed"))
                            {
                                if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                                {
                                    Hob9CNC_TeethShape = SelectedOrderCNCHob9.TeethShape;
                                    Hob9CNC_TeethQty = SelectedOrderCNCHob9.TeethQuantity;
                                    Hob9CNC_Diameter = SelectedOrderCNCHob9.Diameter_d;
                                    //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                    var poToDelete = ListOrdersCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob9.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCHob9.Remove(poToDelete);
                                        JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob9.json");
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWCNCHob9.json");
                                        Task.Delay(1000).Wait();
                                        //2019-10-15 Start add: enable timer again
                                        timerCNCHob.Start();
                                        //2019-10-15 End add
                                    }
                                    //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    btnClear.click();
                                    //2019-10-04 End add

                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable BWHobLoading while the printer is running
                                    IsGettingDataCNC = true;
                                    CNCHob9_IsGettingData = true;
                                    CNCHob9_IsPrinting = false;
                                    //2019-10-07 End add
                                    pageNumberCNCHob = 0;//quay lại từ đầu
                                }
                                else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                                {
                                    Hob10CNC_TeethShape = SelectedOrderCNCHob10.TeethShape;
                                    Hob10CNC_TeethQty = SelectedOrderCNCHob10.TeethQuantity;
                                    Hob10CNC_Diameter = SelectedOrderCNCHob10.Diameter_d;
                                    //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                    var poToDelete = ListOrdersCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob10.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCHob10.Remove(poToDelete);
                                        JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob10.json");
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWCNCHob10.json");
                                        Task.Delay(1000).Wait();
                                        //2019-10-15 Start add: enable timer again
                                        timerCNCHob.Start();
                                        //2019-10-15 End add
                                    }
                                    //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 0
                                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    btnClear.click();
                                    //2019-10-04 End add

                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable BWHobLoading while the printer is running
                                    IsGettingDataCNC = true;
                                    CNCHob10_IsGettingData = true;
                                    CNCHob10_IsPrinting = false;
                                    //2019-10-07 End add
                                    pageNumberCNCHob = 0;//quay lại từ đầu
                                }


                            }
                            else if (status.ToLower().Contains("print target"))
                            {

                                if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob9.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCHob9.Remove(poToDelete);

                                    }
                                    //2019-10-11 Start add: print message: Order has been printed
                                    CNC_Hob9Status = "The selected order has been printed from other machine!";
                                    //2019-10-07 End add
                                    CNCHob9_IsGettingData = true;
                                    CNCHob9_IsPrinting = false;
                                }
                                else if (CNCHob10_IsPrinting) //nếu CNCHob10 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob10.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCHob10.Remove(poToDelete);

                                    }
                                    //2019-10-11 Start add: print message: Order has been printed
                                    CNC_Hob10Status = "The selected order has been printed from other machine!";
                                    //2019-10-07 End add
                                    CNCHob10_IsGettingData = true;
                                    CNCHob10_IsPrinting = false;
                                }



                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add


                                //Enable btnGetData
                                IsGettingDataCNC = true;

                                //2019-10-07 End add
                                pageNumberCNCHob = 0;//quay lại từ đầu
                            }
                            else
                            {
                                //máy in hết giấy, lỗi data, etc..
                                //quay về trang in
                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCHobLoading if getting error
                                IsGettingDataCNC = true;
                                if (CNCHob9_IsPrinting) //nếu CNCHob9 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersCNCHob9.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob9.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCHob9.Remove(poToDelete);

                                    }
                                    //2019-10-11 Start add: print message: Order has been printed
                                    CNC_Hob9Status = "The selected order has been printed from other machine!";
                                    //2019-10-07 End add
                                    CNCHob9_IsGettingData = true;
                                    CNCHob9_IsPrinting = false;
                                }
                                else if (CNCHob10_IsPrinting) //nếu BWHob12 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersCNCHob10.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob10.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCHob10.Remove(poToDelete);

                                    }
                                    //2019-10-11 Start add: print message: Order has been printed
                                    CNC_Hob10Status = "The selected order has been printed from other machine!";
                                    //2019-10-07 End add
                                    CNCHob10_IsGettingData = true;
                                    CNCHob10_IsPrinting = false;
                                }


                                //2019-10-07 End add

                                pageNumberCNCHob = 0;
                            }

                        }

                    }

                }
                else { }

            }
        }
        //KEY-CNC-3
        private void wbCNCNoHob_LoadCompleted(object sender, NavigationEventArgs e)
        {
            isFirstNavigationCNCNoHob = false;
            WebBrowser wb = sender as WebBrowser;
            mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;
            if (isFirstLoadCNCNoHob)
            {

                //get user login info in App.config
                string user = ConfigurationManager.AppSettings["user"];
                string pass = ConfigurationManager.AppSettings["pass"];

                //login
                mshtml.IHTMLElement userInput = document.getElementById("txtUserId");
                mshtml.IHTMLElement passInput = document.getElementById("txtPassword");
                mshtml.IHTMLElement btnLogin = document.getElementById("cmdLogin");
                userInput.setAttribute("value", user);
                passInput.setAttribute("value", pass);
                btnLogin.click();
                isFirstLoadCNCNoHob = false;
            }
            else
            {

                if (pageNumberCNCNoHob == 0)
                {
                    //nếu trang tải xong
                    if (wb.IsLoaded)
                    {
                        //check session time out
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadCNCNoHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();
                        }
                        else
                        {
                            //still on the session
                            var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                            var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                            script.text = @" 
                                   window.onload = function(){
                                        //navigate 'Manufacturing Instructions Print' menu
                                        window.location = '/manufa/PrintViewInstruction.aspx';
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);
                            HideScriptErrors(wb, true);
                            pageNumberCNCNoHob = 1;
                        }
                    }

                }
                else if (pageNumberCNCNoHob == 1)
                {
                    //check session time out
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if CNCNoHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _CNCNoHob_IsOutOfSession = true;
                        //2019-10-11 End add

                        //navigate to login page
                        isFirstLoadCNCNoHob = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();

                    }
                    else
                    {
                        //2019-10-11 Start add
                        if (_CNCNoHob_IsOutOfSession)
                        {
                            _CNCNoHob_IsOutOfSession = false;//reset to default value
                            if (CNCNoHob_IsPrinting) //nếu CNCNoHob đang in
                            {
                                SelectedOrderCNCNoHob = ListOrdersCNCNoHob[0];
                            }



                            //onclick btnClear to navigate to page 2
                            pageNumberCNCNoHob = 2;
                            var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                            btnClear.click();
                        }
                        //2019-10-11 End add
                        else
                        {
                            var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                            var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                            script.text = @" 
                                   window.onload = function(){                                                                              
                                        //set selected status option for the order that not printed yet
                                        var printMode = document.getElementById('ContentPlaceHolder1_rdoSelectOne');
                                        printMode.checked = true;        
                                          
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);
                            pageNumberCNCNoHob = 2;

                        }

                    }

                }
                else if (pageNumberCNCNoHob == 2)
                {

                    //Disable btnGetData, enable CNCHobLoading while the printer is running   
                    if (CNCNoHob_IsPrinting) //nếu CNCNoHob đang in
                    {
                        CNCNoHob_IsGettingData = false;
                    }


                    if (wb.IsLoaded)
                    {

                        if (CheckSessionExpired(wb))
                        {
                            //2019-10-11 Start add: if CNCNoHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                            _CNCNoHob_IsOutOfSession = true;
                            //2019-10-11 End add

                            //navigate to login page
                            isFirstLoadCNCNoHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {


                            //điền PO (manufa NO) và "Search" đơn để in
                            if (CNCNoHob_IsPrinting) //nếu CNCHob9 đang in
                            {
                                document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                                pageNumberCNCNoHob = 3;
                            }

                        }
                    }
                }
                else if (pageNumberCNCNoHob == 3)
                {
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if CNCNoHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _CNCNoHob_IsOutOfSession = true;
                        //2019-10-11 End add

                        //navigate to login page
                        isFirstLoadCNCNoHob = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();

                    }
                    else
                    {

                        var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                        var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                        script.text = @" 
                                   window.onload = function(){                
                                                                                    
                                        var orderToPrint = document.getElementById('ContentPlaceHolder1_grdPrintList_chkOutput_0');
                                        var btnPrint = document.getElementById('ContentPlaceHolder1_cmdPrint');
                                            
                                        if(orderToPrint != null)
                                        {   
                                            //check vào PO cần in
                                            orderToPrint.checked = 'checked';   
                                            //chọn máy in, 4 = pulley
                                            var printer = document.getElementById('ContentPlaceHolder1_drpPrinter');
                                            printer.selectedIndex = 4
                                            //in 
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },2000);
                                            
                                            
                                        }
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        pageNumberCNCNoHob = 4; //chuyển trang sau khi in, lấy status
                    }


                }
                else if (pageNumberCNCNoHob == 4)
                {
                    if (wb.IsLoaded)
                    {
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadCNCNoHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();

                        }
                        else
                        {
                            //lấy status sau khi in
                            string status = document.getElementById("ContentPlaceHolder1_AlertArea").innerHTML;
                            //nếu in thành công
                            if (status.ToLower().Contains("completed"))
                            {
                                if (CNCNoHob_IsPrinting) //nếu CNCNoHob đang in
                                {
                                    NoHobCNC_TeethShape = SelectedOrderCNCNoHob.TeethShape;
                                    NoHobCNC_TeethQty = SelectedOrderCNCNoHob.TeethQuantity;

                                    //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                    var poToDelete = ListOrdersCNCNoHob.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCNoHob.Remove(poToDelete);
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob9.json");
                                        //JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWCNCHob9.json");
                                        Task.Delay(1000).Wait();
                                        //2019-10-15 Start add: enable timer again
                                        //timerCNCHob.Start();
                                        //2019-10-15 End add
                                    }
                                    //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                    btnClear.click();
                                    //2019-10-04 End add

                                    //2019-10-07 Start add
                                    //Enable btnGetData, disable CNCNoHobLoading while the printer is running
                                    IsGettingDataCNC = true;
                                    CNCNoHob_IsGettingData = true;
                                    CNCNoHob_IsPrinting = false;
                                    //2019-10-07 End add
                                    pageNumberCNCNoHob = 0;//quay lại từ đầu
                                }



                            }
                            else if (status.ToLower().Contains("print target"))
                            {

                                if (CNCNoHob_IsPrinting) //nếu CNCHob9 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersCNCNoHob.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCNoHob.Remove(poToDelete);

                                    }
                                    //2019-10-11 Start add: print message: Order has been printed
                                    CNC_NoHobStatus = "The selected order has been printed from other machine!";
                                    //2019-10-07 End add
                                    CNCNoHob_IsGettingData = true;
                                    CNCNoHob_IsPrinting = false;
                                }




                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add


                                //Enable btnGetData
                                IsGettingDataCNC = true;

                                //2019-10-07 End add
                                pageNumberCNCNoHob = 0;//quay lại từ đầu
                            }
                            else
                            {
                                //máy in hết giấy, lỗi data, etc..
                                //quay về trang in
                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCHobLoading if getting error
                                IsGettingDataCNC = true;
                                if (CNCNoHob_IsPrinting) //nếu CNCHob9 đang in
                                {
                                    //xóa PO đã in thành công khỏi list
                                    var poToDelete = ListOrdersCNCNoHob.FirstOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                    if (poToDelete != null)
                                    {
                                        ListOrdersCNCNoHob.Remove(poToDelete);

                                    }
                                    //2019-10-11 Start add: print message: Order has been printed
                                    CNC_NoHobStatus = "The selected order has been printed from other machine!";
                                    //2019-10-07 End add
                                    CNCNoHob_IsGettingData = true;
                                    CNCNoHob_IsPrinting = false;
                                }



                                //2019-10-07 End add

                                pageNumberCNCNoHob = 0;
                            }

                        }

                    }

                }
                else { }

            }
        }
        #endregion <<COMMAND-FUNCTIONS>>

        #region <<BW-COMMAND-FUNCTIONS>>

        //2019-12-03 Start add
        private void ToggleOnOff_BWCNCHob9(object obj)
        {
            if(obj != null)
            {
                var values = obj as object[];               
                ToggleButton tg = values[0] as ToggleButton;
                if(tg.IsChecked == true)
                {
                    //stop timerBWHob để tạm dừng đồng bộ info từ MYSQL 
                    timerBWHob.Stop();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob9BWCNC_TeethShape_IsEnabled = true;
                    Hob9BWCNC_TeethQty_IsEnabled = true;
                    Hob9BWCNC_Diameter_IsEnabled = true;
                    Hob9BWCNC_GlobalCode_IsEnabled = true;
                }
                else
                {
                    //lưu thông tin vào MYSQL
                    try
                    {
                        int result = OrderBLL.Instance.EditBWCNCHob9("",Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_DMat);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Cập nhật thông tin Teeth thất bại, vui lòng kiểm tra lại!\n"+ex.Message,"Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    //khởi động lại timer để đồng bộ info cho BWCNCHob9
                    timerBWHob.Start();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob9BWCNC_TeethShape_IsEnabled = false;
                    Hob9BWCNC_TeethQty_IsEnabled = false;
                    Hob9BWCNC_Diameter_IsEnabled = false;
                    Hob9BWCNC_GlobalCode_IsEnabled = false;
                }
            }
        }
        private void ToggleOnOff_BWCNCHob10(object obj)
        {
            if (obj != null)
            {
                var values = obj as object[];
                ToggleButton tg = values[0] as ToggleButton;
                if (tg.IsChecked == true)
                {
                    //stop timerBWHob để tạm dừng đồng bộ info từ MYSQL 
                    timerBWHob.Stop();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob10BWCNC_TeethShape_IsEnabled = true;
                    Hob10BWCNC_TeethQty_IsEnabled = true;
                    Hob10BWCNC_Diameter_IsEnabled = true;
                    Hob10BWCNC_GlobalCode_IsEnabled = true;
                }
                else
                {
                    //lưu thông tin vào MYSQL
                    try
                    {
                        int result = OrderBLL.Instance.EditBWCNCHob10("", Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_DMat);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Cập nhật thông tin Teeth thất bại, vui lòng kiểm tra lại!\n" + ex.Message, "Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    //khởi động lại timer để đồng bộ info cho BWCNCHob10
                    timerBWHob.Start();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob10BWCNC_TeethShape_IsEnabled = false;
                    Hob10BWCNC_TeethQty_IsEnabled = false;
                    Hob10BWCNC_Diameter_IsEnabled = false;
                    Hob10BWCNC_GlobalCode_IsEnabled = false;
                }
            }
        }
        //2019-12-03 End add
        private async void PrintBWHob11(object parameter)
        {
            if (parameter != null)
            {
                var values = parameter as object[];
                UserControl ucBWHob11 = values[0] as UserControl;
                WebBrowser wb = ucBWHob11.FindName("wbBWHob11") as WebBrowser;
                await Task.Run(() =>
               {
                   wb.Dispatcher.Invoke(() =>
                   {
                       mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;


                       IsGettingDataBW = false;
                       BWHob11_IsPrinting = true;
                       BWHob11_IsGettingData = false;
                       timerBWHob.Stop();
                       SelectedOrderBWHob11 = ListOrdersBWHob11.FirstOrDefault();

                       //2019-11-05 Start add: điền PO vào và tìm trên Manufa, bỏ qua bước clear hệ thống Manufa rồi mới tìm PO
                       var table = document.getElementById("ContentPlaceHolder1_grdPrintList");
                       HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                       doc.LoadHtml(table.outerHTML);
                       var rows = doc.GetElementbyId("ContentPlaceHolder1_grdPrintList").SelectNodes("//*[@id='ContentPlaceHolder1_grdPrintList']/tbody/tr");
                       if (rows.Count < 2)
                       {
                           document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob11.Manufa_Instruction_No);
                           document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                           //điều hướng trang 
                           pageNumberBWHob11 = 3;
                       }
                       else
                       {
                           document.getElementById("ContentPlaceHolder1_cmdPrint").click();
                           //điều hướng trang
                           pageNumberBWHob11 = 3;
                       }

                       //2019-11-05 End add

                       //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, thông báo PO đang in
                       BW_Hob11Status = "Đang in PO: " + SelectedOrderBWHob11.Manufa_Instruction_No;
                       //SendMailNotificationBW();
                       //2019-11-28 End add
                   });


               });

            }
        }

        private bool CanPrintBWHob11(object obj)
        {
            if (ListOrdersBWHob11 != null && ListOrdersBWHob11.Count > 0 && IsLoadingBW == false && BWHob12_IsPrinting == false && BWCNCHob9_IsPrinting == false && BWCNCHob10_IsPrinting == false)
                return true;
            else return false;
        }

        private async void PrintBWHob12(object parameter)
        {
            if (parameter != null)
            {               
                #region 2019-11-08 In chung webbrowser của Hob11
                var values = parameter as object[];
                UserControl ucBWHob11 = values[0] as UserControl;
                WebBrowser wb = ucBWHob11.FindName("wbBWHob11") as WebBrowser;

                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

                        IsGettingDataBW = false;
                        BWHob12_IsPrinting = true;
                        BWHob12_IsGettingData = false;
                        timerBWHob.Stop();
                        SelectedOrderBWHob12 = ListOrdersBWHob12.FirstOrDefault();
                        //2019-11-05 Start lock: đổi cách in
                        ////2019-11-4 Start add: sau khi in xong (còn trạng thái Printed Completed) thì in tiếp ko quay về trang đầu tiên
                        //pageNumberBWHob = 2;
                        ////2019-11-4 End add                
                        //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                        //btnClear.click();
                        //2019-11-05 End lock

                        //2019-11-05 Start add: điền PO vào và tìm trên Manufa, bỏ qua bước clear hệ thống Manufa rồi mới tìm PO
                        var table = document.getElementById("ContentPlaceHolder1_grdPrintList");
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(table.outerHTML);
                        var rows = doc.GetElementbyId("ContentPlaceHolder1_grdPrintList").SelectNodes("//*[@id='ContentPlaceHolder1_grdPrintList']/tbody/tr");
                        if (rows.Count < 2)
                        {
                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob12.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            //điều hướng trang 
                            pageNumberBWHob11 = 3;
                        }
                        else
                        {
                            document.getElementById("ContentPlaceHolder1_cmdPrint").click();
                            //điều hướng trang
                            pageNumberBWHob11 = 3;
                        }

                        //2019-11-05 End add

                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, thông báo in PO
                        BW_Hob12Status = "Đang in PO: " + SelectedOrderBWHob12.Manufa_Instruction_No;
                        //SendMailNotificationBW();
                        //2019-11-28 End add
                    });

                });
                #endregion


            }
        }

        private bool CanPrintBWHob12(object obj)
        {
            if (ListOrdersBWHob12 != null && ListOrdersBWHob12.Count > 0 && IsLoadingBW == false && BWHob11_IsPrinting == false && BWCNCHob9_IsPrinting == false && BWCNCHob10_IsPrinting == false)
                return true;
            else return false;
        }

        private async void PrintBWNoHob(object obj)
        {
            if (obj != null)
            {
                var values = obj as object[];
                Awesomium.Windows.Controls.WebControl wb = values[0] as Awesomium.Windows.Controls.WebControl;
                //UserControl ucBWHob11 = values[0] as UserControl;
                //WebBrowser wb = ucBWHob11.FindName("wbBWHob11") as WebBrowser;


                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        dynamic document = (Awesomium.Core.JSObject)wb.ExecuteJavascriptWithResult("document");
                        //mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

                        IsGettingDataBW = false;
                        BWNoHob_IsPrinting = true;
                        BWNoHob_IsGettingData = false;
                        timerBWHob.Stop();


                        //2019-11-05 Start add: điền PO vào và tìm trên Manufa, bỏ qua bước clear hệ thống Manufa rồi mới tìm PO
                        if (SelectedOrderBWNoHob.LineAB.Equals("BW"))
                        {
                            

                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWNoHob.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            //điều hướng trang 
                            pageNumberBWNoHob = 3;

                        }

                        //2019-11-05 End add

                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, đang in PO
                        BW_NoHobStatus = "Đang in PO: " + SelectedOrderBWNoHob.Manufa_Instruction_No;
                        //SendMailNotificationBW();
                        //2019-11-28 End add
                    });

                });


            }
           
        }

        private bool CanPrintBWNoHob(object obj)
        {
            if (SelectedOrderBWNoHob != null && IsLoadingBW == false)
                return true;
            else return false;
        }

        private async void PrintBWF1(object obj)
        {
            #region in chung browser BWNoHob
            if (obj != null)
            {
                var values = obj as object[];
                Awesomium.Windows.Controls.WebControl wb = values[0] as Awesomium.Windows.Controls.WebControl;
                //UserControl ucBWHob11 = values[0] as UserControl;
                //WebBrowser wb = ucBWHob11.FindName("wbBWHob11") as WebBrowser;
                //wb.LoadingFrameComplete += wbBWNoHob_LoadedFrameComplete;

                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        dynamic document = (Awesomium.Core.JSObject)wb.ExecuteJavascriptWithResult("document");
                        //mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

                        IsGettingDataBW = false;
                        BWF1_IsPrinting = true;
                        BWF1_IsGettingData = false;
                        timerBWHob.Stop();

                        
                        //2019-11-05 Start add: điền PO vào và tìm trên Manufa, bỏ qua bước clear hệ thống Manufa rồi mới tìm PO
                        if (SelectedOrderBWF1 != null && SelectedOrderBWF1.LineAB.Equals("BW"))
                        {


                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWF1.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            //điều hướng trang 
                            pageNumberBWNoHob = 3;

                        }
                     

                        //2019-11-05 End add

                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, đang in PO
                        BW_F1Status = "Đang in PO: " + SelectedOrderBWF1.Manufa_Instruction_No;
                        //SendMailNotificationBW();
                        //2019-11-28 End add
                    });

                });


            }
            #endregion

        }
        private bool CanPrintBWF1(object obj)
        {
            if (ListOrdersBWF1!= null && ListOrdersBWF1.Count > 0 && IsLoadingBW == false && SelectedOrderBWF1 != null)
                return true;
            else return false;
        }

        private async void PrintBWCNCHob9(object parameter)
        {
            if (parameter != null)
            {
                var values = parameter as object[];
                UserControl ucBWHob11 = values[0] as UserControl;
                WebBrowser wb = ucBWHob11.FindName("wbBWHob11") as WebBrowser;
                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;


                        IsGettingDataBW = false;
                        BWCNCHob9_IsPrinting = true;
                        BWCNCHob9_IsGettingData = false;
                        timerBWHob.Stop();
                        SelectedOrderBWCNCHob9 = ListOrdersBWCNCHob9.FirstOrDefault();

                        //2019-11-05 Start lock: đổi cách in
                        ////2019-11-4 Start add: sau khi in xong (còn trạng thái Printed Completed) thì in tiếp ko quay về trang đầu tiên
                        //pageNumberBWHob = 2;
                        ////2019-11-4 End add                
                        //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                        //btnClear.click();
                        //2019-11-05 End lock

                        //2019-11-05 Start add: điền PO vào và tìm trên Manufa, bỏ qua bước clear hệ thống Manufa rồi mới tìm PO
                        var table = document.getElementById("ContentPlaceHolder1_grdPrintList");
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(table.outerHTML);
                        var rows = doc.GetElementbyId("ContentPlaceHolder1_grdPrintList").SelectNodes("//*[@id='ContentPlaceHolder1_grdPrintList']/tbody/tr");
                        if (rows.Count < 2)
                        {
                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWCNCHob9.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            //điều hướng trang 
                            pageNumberBWHob11 = 3;
                        }
                        else
                        {
                            document.getElementById("ContentPlaceHolder1_cmdPrint").click();
                            //điều hướng trang
                            pageNumberBWHob11 = 3;
                        }

                        //2019-11-05 End add

                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, thông báo PO đang in
                        BWCNC_Hob9Status = "Đang in PO: " + SelectedOrderBWCNCHob9.Manufa_Instruction_No;
                        //SendMailNotificationBW();
                        //2019-11-28 End add
                    });


                });

            }
        }

        private bool CanPrintBWCNCHob9(object obj)
        {
            if (ListOrdersBWCNCHob9 != null && ListOrdersBWCNCHob9.Count > 0 && IsLoadingBW == false && BWHob11_IsPrinting == false && BWHob12_IsPrinting == false && BWCNCHob10_IsPrinting == false && ListOrdersBWCNCHob9 != null && ListOrdersBWCNCHob9.Count > 0)
            {
                SelectedOrderBWCNCHob9 = ListOrdersBWCNCHob9[0];
                if (!SelectedOrderBWCNCHob9.LineAB.Equals("CNC"))
                    return true; //&& !SelectedOrderBWCNCHob9.LineAB.Equals("CNC")
                else return false;
            }
            else return false;
        }
        private async void PrintBWCNCHob10(object parameter)
        {           
            if (parameter != null)
            {
                var values = parameter as object[];
                UserControl ucBWHob11 = values[0] as UserControl;
                WebBrowser wb = ucBWHob11.FindName("wbBWHob11") as WebBrowser;
                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;


                        IsGettingDataBW = false;
                        BWCNCHob10_IsPrinting = true;
                        BWCNCHob10_IsGettingData = false;
                        timerBWHob.Stop();
                        SelectedOrderBWCNCHob10 = ListOrdersBWCNCHob10.FirstOrDefault();

                        //2019-11-05 Start lock: đổi cách in
                        ////2019-11-4 Start add: sau khi in xong (còn trạng thái Printed Completed) thì in tiếp ko quay về trang đầu tiên
                        //pageNumberBWHob = 2;
                        ////2019-11-4 End add                
                        //var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                        //btnClear.click();
                        //2019-11-05 End lock

                        //2019-11-05 Start add: điền PO vào và tìm trên Manufa, bỏ qua bước clear hệ thống Manufa rồi mới tìm PO
                        var table = document.getElementById("ContentPlaceHolder1_grdPrintList");
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(table.outerHTML);
                        var rows = doc.GetElementbyId("ContentPlaceHolder1_grdPrintList").SelectNodes("//*[@id='ContentPlaceHolder1_grdPrintList']/tbody/tr");
                        if (rows.Count < 2)
                        {
                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWCNCHob10.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            //điều hướng trang 
                            pageNumberBWHob11 = 3;
                        }
                        else
                        {
                            document.getElementById("ContentPlaceHolder1_cmdPrint").click();
                            //điều hướng trang
                            pageNumberBWHob11 = 3;
                        }

                        //2019-11-05 End add

                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, thông báo PO đang in
                        BWCNC_Hob10Status = "Đang in PO: " + SelectedOrderBWCNCHob10.Manufa_Instruction_No;
                        //SendMailNotificationBW();
                        //2019-11-28 End add
                    });


                });

            }
        }

        private bool CanPrintBWCNCHob10(object obj)
        {
            if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count > 0 && IsLoadingBW == false && BWHob11_IsPrinting == false && BWHob12_IsPrinting == false && BWCNCHob9_IsPrinting == false && ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count > 0)
            {
                SelectedOrderBWCNCHob10 = ListOrdersBWCNCHob10[0];
                if (!SelectedOrderBWCNCHob10.LineAB.Equals("CNC"))
                    return true; //&& !SelectedOrderBWCNCHob10.LineAB.Equals("CNC")
                else return false;
            }
            else return false;
        }

        #endregion <<BW-COMMAND-FUNCTIONS>>

        #region <<CNC-COMMAND-FUNCTIONS>>
        //2019-12-03 Start add
        private void ToggleOnOff_CNCHob9(object obj)
        {
            if (obj != null)
            {
                var values = obj as object[];
                ToggleButton tg = values[0] as ToggleButton;
                if (tg.IsChecked == true)
                {
                    //stop timerCNCHob để tạm dừng đồng bộ info từ MYSQL 
                    timerCNCHob.Stop();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob9CNC_TeethShape_IsEnabled = true;
                    Hob9CNC_TeethQty_IsEnabled = true;
                    Hob9CNC_Diameter_IsEnabled = true;

                }
                else
                {
                    //lưu thông tin vào MYSQL
                    try
                    {
                        int result = OrderBLL.Instance.EditBWCNCHob9("", Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter, Hob9CNC_DMat);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Cập nhật thông tin Teeth thất bại, vui lòng kiểm tra lại!\n" + ex.Message, "Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    //khởi động lại timer để đồng bộ info cho CNCHob9
                    timerCNCHob.Start();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob9CNC_TeethShape_IsEnabled = false;
                    Hob9CNC_TeethQty_IsEnabled = false;
                    Hob9CNC_Diameter_IsEnabled = false;
                }
            }
        }
        private void ToggleOnOff_CNCHob10(object obj)
        {
            if (obj != null)
            {
                var values = obj as object[];
                ToggleButton tg = values[0] as ToggleButton;
                if (tg.IsChecked == true)
                {
                    //stop timerCNCHob để tạm dừng đồng bộ info từ MYSQL 
                    timerCNCHob.Stop();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob10CNC_TeethShape_IsEnabled = true;
                    Hob10CNC_TeethQty_IsEnabled = true;
                    Hob10CNC_Diameter_IsEnabled = true;

                }
                else
                {
                    //lưu thông tin vào MYSQL
                    try
                    {
                        int result = OrderBLL.Instance.EditBWCNCHob10("", Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter, Hob10CNC_DMat);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Cập nhật thông tin Teeth thất bại, vui lòng kiểm tra lại!\n" + ex.Message, "Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    //khởi động lại timer để đồng bộ info cho CNCHob9
                    timerCNCHob.Start();
                    //khóa các thông số TeethShape, TeethQty, Diameter_d
                    Hob10CNC_TeethShape_IsEnabled = false;
                    Hob10CNC_TeethQty_IsEnabled = false;
                    Hob10CNC_Diameter_IsEnabled = false;
                }
            }
        }
        //2019-12-03 End add
        private async void PrintCNCHob9(object parameter)
        {
            if (parameter != null)
            {
                

                var values = parameter as object[];
                Awesomium.Windows.Controls.WebControl wb = values[0] as Awesomium.Windows.Controls.WebControl;
                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        dynamic document = (Awesomium.Core.JSObject)wb.ExecuteJavascriptWithResult("document");
                        IsGettingDataCNC = false;
                        CNCHob9_IsPrinting = true;
                        CNCHob9_IsGettingData = false;
                        timerCNCHob.Stop();

                        SelectedOrderCNCHob9 = ListOrdersCNCHob9.FirstOrDefault();
                        using (document)
                        {
                            wb.ExecuteJavascriptWithResult(@"
                            document.getElementById('ContentPlaceHolder1_txtAufnr').setAttribute('value', " + SelectedOrderCNCHob9.Manufa_Instruction_No + "); " +
                            "document.getElementById('ContentPlaceHolder1_cmbSelect').click();");
                            //document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCHob9.Manufa_Instruction_No);
                            //document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            //điều hướng trang 
                            pageNumberCNCHob = 3;
                        }
                        
                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, đang in PO
                        CNC_Hob9Status = "Đang in PO: " + SelectedOrderCNCHob9.Manufa_Instruction_No;
                        //SendMailNotificationCNC();
                        //2019-11-28 End add

                    });
                });
            }
        }

        private bool CanPrintCNCHob9(object obj)
        {
            if (ListOrdersCNCHob9 != null && ListOrdersCNCHob9.Count > 0 && IsLoadingCNC == false && _CNCHob10_IsPrinting == false)
                return true;
            else return false;
        }

        private async void PrintCNCHob10(object parameter)
        {
            if (parameter != null)
            {

                var values = parameter as object[];
                Awesomium.Windows.Controls.WebControl wb = values[0] as Awesomium.Windows.Controls.WebControl;
                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        dynamic document = (Awesomium.Core.JSObject)wb.ExecuteJavascriptWithResult("document");
                        IsGettingDataCNC = false;
                        CNCHob10_IsPrinting = true;
                        CNCHob10_IsGettingData = false;
                        timerCNCHob.Stop();
                        SelectedOrderCNCHob10 = ListOrdersCNCHob10.FirstOrDefault();
                        
                        document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCHob10.Manufa_Instruction_No);
                        document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                        //điều hướng trang 
                        pageNumberCNCHob = 3;


                        //2019-11-05 End add

                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, đang in PO
                        CNC_Hob10Status = "Đang in PO: " + SelectedOrderCNCHob10.Manufa_Instruction_No;
                        //SendMailNotificationCNC();
                        //2019-11-28 End add

                    });
                });
            }
        }

        private bool CanPrintCNCHob10(object obj)
        {
            if (ListOrdersCNCHob10 != null && ListOrdersCNCHob10.Count > 0 && IsLoadingCNC == false && _CNCHob9_IsPrinting == false)
                return true;
            else return false;
        }
        private async void PrintCNCNoHob(object parameter)
        {
            if (parameter != null)
            {
                
                var values = parameter as object[];
                Awesomium.Windows.Controls.WebControl wb = values[0] as Awesomium.Windows.Controls.WebControl;
                await Task.Run(() =>
                {
                    wb.Dispatcher.Invoke(() =>
                    {
                        dynamic document = (Awesomium.Core.JSObject)wb.ExecuteJavascriptWithResult("document");
                        IsGettingDataCNC = false;
                        CNCNoHob_IsPrinting = true;
                        CNCNoHob_IsGettingData = false;
                        timerCNCHob.Stop();                        

                        document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCNoHob.Manufa_Instruction_No);
                        document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                        //điều hướng trang 
                        pageNumberCNCHob = 3;


                        //2019-11-05 End add

                        //2019-11-28 Start add: gửi mail PC nếu Order dưới 6, đang in PO
                        CNC_NoHobStatus = "Đang in PO: " + SelectedOrderCNCNoHob.Manufa_Instruction_No;
                        //SendMailNotificationCNC();
                        //2019-11-28 End add

                    });
                });
            }
        }

        private bool CanPrintCNCNoHob(object obj)
        {
            if (SelectedOrderCNCNoHob != null && IsLoadingCNC == false)
                return true;
            else return false;
        }
        #endregion <<CNC-COMMAND-FUNCTIONS>>

        private void KillChromeDriver()
        {
            //foreach (var process in Process.GetProcessesByName("chromedriver"))
            //{
            //    process.Kill();
            //}
        }

        private bool CheckSessionExpired(WebBrowser browser)
        {
            if (browser.Source.AbsoluteUri.Equals("http://10.4.24.111:8080/error/SessionTimeOut.aspx"))
                return true;
            else return false;
        }

        private bool CheckSessionExpiredAwesomiumBrowser(Awesomium.Windows.Controls.WebControl browser)
        {
            if (browser.Source.AbsoluteUri.Equals("http://10.4.24.111:8080/error/SessionTimeOut.aspx"))
                return true;
            else return false;
        }

        public void HideScriptErrors(WebBrowser wb, bool hide)
        {
            var fiComWebBrowser = typeof(WebBrowser).GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fiComWebBrowser == null) return;
            var objComWebBrowser = fiComWebBrowser.GetValue(wb);
            if (objComWebBrowser == null)
            {
                wb.Loaded += (o, s) => HideScriptErrors(wb, hide); //In case we are to early
                return;
            }
            objComWebBrowser.GetType().InvokeMember("Silent", BindingFlags.SetProperty, null, objComWebBrowser, new object[] { hide });
        }

        private ObservableCollection<OrderModel> GetBWOrdersFromExcel(string excelFile)
        {
            ObservableCollection<OrderModel> orders = new ObservableCollection<OrderModel>();
            try
            {
                FileInfo file = new FileInfo(excelFile);
                if (File.Exists(excelFile))
                {
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        //get first WorkSheet
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                        int columnCount = worksheet.Dimension.End.Column;
                        int rowCount = worksheet.Dimension.End.Row;
                        if (worksheet != null)
                        {
                            DateTime validDatetime;
                            //start from row 2nd skip first row for headers
                            for (int i = 2; i <= rowCount; i++)
                            {
                                orders.Add(new OrderModel()
                                {
                                    No = worksheet.Cells[i, 1].Value.ToString() ?? "",
                                    Sales_Order_No = worksheet.Cells[i, 2].Value.ToString() ?? "",
                                    Pier_Instruction_No = worksheet.Cells[i, 3].Value.ToString() ?? "",
                                    Manufa_Instruction_No = worksheet.Cells[i, 4].Value.ToString() ?? "",
                                    Global_Code = worksheet.Cells[i, 5].Value.ToString() ?? "",
                                    Customers = worksheet.Cells[i, 6].Value.ToString() ?? "",
                                    Item_Name = worksheet.Cells[i, 7].Value.ToString() ?? "",
                                    MC = worksheet.Cells[i, 8].Value.ToString() ?? "",
                                    //2019-10-03 Start delete
                                    //Received_Date = !string.IsNullOrEmpty(worksheet.Cells[i, 9].Value.ToString()) ? Convert.ToDateTime(worksheet.Cells[i, 9].Value) : DateTime.Now,
                                    //2019-10-03 End delete
                                    Received_Date = DateTime.TryParse((string)worksheet.Cells[i, 9].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    Factory_Ship_Date = DateTime.TryParse((string)worksheet.Cells[i, 10].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    //mai chỉnh lại kiểu ngày mặc định
                                    Number_of_Orders = worksheet.Cells[i, 11].Value.ToString() ?? "",
                                    Number_of_Available_Instructions = worksheet.Cells[i, 12].Value.ToString() ?? "",
                                    Number_of_Repairs = worksheet.Cells[i, 13].Value.ToString() ?? "",
                                    Number_of_Instructions = worksheet.Cells[i, 14].Value.ToString() ?? "",
                                    Line = worksheet.Cells[i, 15].Value.ToString() ?? "",
                                    PayWard = worksheet.Cells[i, 16].Value.ToString() ?? "",
                                    Major = worksheet.Cells[i, 17].Value.ToString() ?? "",
                                    Special_Orders = worksheet.Cells[i, 18].Value.ToString() ?? "",
                                    Method = worksheet.Cells[i, 19].Value.ToString() ?? "",
                                    Destination = worksheet.Cells[i, 20].Value.ToString() ?? "",
                                    Instructions_Print_date = DateTime.TryParse((string)worksheet.Cells[i, 21].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    Latest_progress = worksheet.Cells[i, 22].Value.ToString() ?? "",
                                    Tack_Label_Output_Date = worksheet.Cells[i, 23].Value.ToString() ?? "",
                                    Completion_Instruction_Date = DateTime.TryParse((string)worksheet.Cells[i, 24].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    Re_print_Count = worksheet.Cells[i, 25].Value.ToString() ?? "",
                                    Latest_issue_time = worksheet.Cells[i, 26].Value.ToString() ?? "",
                                    History1 = worksheet.Cells[i, 27].Value.ToString() ?? "",
                                    History2 = worksheet.Cells[i, 28].Value.ToString() ?? "",
                                    History3 = worksheet.Cells[i, 29].Value.ToString() ?? "",
                                    History4 = worksheet.Cells[i, 30].Value.ToString() ?? "",
                                    Material_code1 = worksheet.Cells[i, 31].Value.ToString() ?? "",
                                    Material_text1 = worksheet.Cells[i, 32].Value.ToString() ?? "",
                                    Amount_used1 = worksheet.Cells[i, 33].Value.ToString() ?? "",
                                    Unit1 = worksheet.Cells[i, 34].Value.ToString() ?? "",
                                    Material_code2 = worksheet.Cells[i, 35].Value.ToString() ?? "",
                                    Material_text2 = worksheet.Cells[i, 36].Value.ToString() ?? "",
                                    Amount_used2 = worksheet.Cells[i, 37].Value.ToString() ?? "",
                                    Unit2 = worksheet.Cells[i, 38].Value.ToString() ?? "",
                                    Material_code3 = worksheet.Cells[i, 39].Value.ToString() ?? "",
                                    Material_text3 = worksheet.Cells[i, 40].Value.ToString() ?? "",
                                    Amount_used3 = worksheet.Cells[i, 41].Value.ToString() ?? "",
                                    Unit3 = worksheet.Cells[i, 42].Value.ToString() ?? "",
                                    Material_code4 = worksheet.Cells[i, 43].Value.ToString() ?? "",
                                    Material_text4 = worksheet.Cells[i, 44].Value.ToString() ?? "",
                                    Amount_used4 = worksheet.Cells[i, 45].Value.ToString() ?? "",
                                    Unit4 = worksheet.Cells[i, 46].Value.ToString() ?? "",
                                    Material_code5 = worksheet.Cells[i, 47].Value.ToString() ?? "",
                                    Material_text5 = worksheet.Cells[i, 48].Value.ToString() ?? "",
                                    Amount_used5 = worksheet.Cells[i, 49].Value.ToString() ?? "",
                                    Unit5 = worksheet.Cells[i, 50].Value.ToString() ?? "",
                                    Material_code6 = worksheet.Cells[i, 51].Value.ToString() ?? "",
                                    Material_text6 = worksheet.Cells[i, 52].Value.ToString() ?? "",
                                    Amount_used6 = worksheet.Cells[i, 53].Value.ToString() ?? "",
                                    Unit6 = worksheet.Cells[i, 54].Value.ToString() ?? "",
                                    Material_code7 = worksheet.Cells[i, 55].Value.ToString() ?? "",
                                    Material_text7 = worksheet.Cells[i, 56].Value.ToString() ?? "",
                                    Amount_used7 = worksheet.Cells[i, 57].Value.ToString() ?? "",
                                    Unit7 = worksheet.Cells[i, 58].Value.ToString() ?? "",
                                    Material_code8 = worksheet.Cells[i, 59].Value.ToString() ?? "",
                                    Material_text8 = worksheet.Cells[i, 60].Value.ToString() ?? "",
                                    Amount_used8 = worksheet.Cells[i, 61].Value.ToString() ?? "",
                                    Unit8 = worksheet.Cells[i, 62].Value.ToString() ?? "",
                                    Inner_Code = worksheet.Cells[i, 63].Value.ToString() ?? "",
                                    Classify_Code = worksheet.Cells[i, 64].Value.ToString() ?? "",
                                });
                            }
                        }

                    }
                }

                return orders;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private ObservableCollection<OrderModel> GetCNCOrdersFromExcel(string excelFile)
        {
            ObservableCollection<OrderModel> orders = new ObservableCollection<OrderModel>();
            try
            {
                FileInfo file = new FileInfo(excelFile);
                if (File.Exists(excelFile))
                {
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        //get first WorkSheet
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                        int columnCount = worksheet.Dimension.End.Column;
                        int rowCount = worksheet.Dimension.End.Row;
                        if (worksheet != null)
                        {
                            DateTime validDatetime;
                            //start from row 2nd skip first row for headers
                            for (int i = 2; i <= rowCount; i++)
                            {
                                orders.Add(new OrderModel()
                                {
                                    No = worksheet.Cells[i, 1].Value.ToString() ?? "",
                                    Sales_Order_No = worksheet.Cells[i, 2].Value.ToString() ?? "",
                                    Pier_Instruction_No = worksheet.Cells[i, 3].Value.ToString() ?? "",
                                    Manufa_Instruction_No = worksheet.Cells[i, 4].Value.ToString() ?? "",
                                    Global_Code = worksheet.Cells[i, 5].Value.ToString() ?? "",
                                    Customers = worksheet.Cells[i, 6].Value.ToString() ?? "",
                                    Item_Name = worksheet.Cells[i, 7].Value.ToString() ?? "",
                                    MC = worksheet.Cells[i, 8].Value.ToString() ?? "",
                                    //2019-10-03 Start delete
                                    //Received_Date = !string.IsNullOrEmpty(worksheet.Cells[i, 9].Value.ToString()) ? Convert.ToDateTime(worksheet.Cells[i, 9].Value) : DateTime.Now,
                                    //2019-10-03 End delete
                                    Received_Date = DateTime.TryParse((string)worksheet.Cells[i, 9].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    Factory_Ship_Date = DateTime.TryParse((string)worksheet.Cells[i, 10].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    //mai chỉnh lại kiểu ngày mặc định
                                    Number_of_Orders = worksheet.Cells[i, 11].Value.ToString() ?? "",
                                    Number_of_Available_Instructions = worksheet.Cells[i, 12].Value.ToString() ?? "",
                                    Number_of_Repairs = worksheet.Cells[i, 13].Value.ToString() ?? "",
                                    Number_of_Instructions = worksheet.Cells[i, 14].Value.ToString() ?? "",
                                    Line = worksheet.Cells[i, 15].Value.ToString() ?? "",
                                    PayWard = worksheet.Cells[i, 16].Value.ToString() ?? "",
                                    Major = worksheet.Cells[i, 17].Value.ToString() ?? "",
                                    Special_Orders = worksheet.Cells[i, 18].Value.ToString() ?? "",
                                    Method = worksheet.Cells[i, 19].Value.ToString() ?? "",
                                    Destination = worksheet.Cells[i, 20].Value.ToString() ?? "",
                                    Instructions_Print_date = DateTime.TryParse((string)worksheet.Cells[i, 21].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    Latest_progress = worksheet.Cells[i, 22].Value.ToString() ?? "",
                                    Tack_Label_Output_Date = worksheet.Cells[i, 23].Value.ToString() ?? "",
                                    Completion_Instruction_Date = DateTime.TryParse((string)worksheet.Cells[i, 24].Value, out validDatetime) ? validDatetime : (DateTime?)null,
                                    Re_print_Count = worksheet.Cells[i, 25].Value.ToString() ?? "",
                                    Latest_issue_time = worksheet.Cells[i, 26].Value.ToString() ?? "",
                                    History1 = worksheet.Cells[i, 27].Value.ToString() ?? "",
                                    History2 = worksheet.Cells[i, 28].Value.ToString() ?? "",
                                    History3 = worksheet.Cells[i, 29].Value.ToString() ?? "",
                                    History4 = worksheet.Cells[i, 30].Value.ToString() ?? "",
                                    Material_code1 = worksheet.Cells[i, 31].Value.ToString() ?? "",
                                    Material_text1 = worksheet.Cells[i, 32].Value.ToString() ?? "",
                                    Amount_used1 = worksheet.Cells[i, 33].Value.ToString() ?? "",
                                    Unit1 = worksheet.Cells[i, 34].Value.ToString() ?? "",
                                    Material_code2 = worksheet.Cells[i, 35].Value.ToString() ?? "",
                                    Material_text2 = worksheet.Cells[i, 36].Value.ToString() ?? "",
                                    Amount_used2 = worksheet.Cells[i, 37].Value.ToString() ?? "",
                                    Unit2 = worksheet.Cells[i, 38].Value.ToString() ?? "",
                                    Material_code3 = worksheet.Cells[i, 39].Value.ToString() ?? "",
                                    Material_text3 = worksheet.Cells[i, 40].Value.ToString() ?? "",
                                    Amount_used3 = worksheet.Cells[i, 41].Value.ToString() ?? "",
                                    Unit3 = worksheet.Cells[i, 42].Value.ToString() ?? "",
                                    Material_code4 = worksheet.Cells[i, 43].Value.ToString() ?? "",
                                    Material_text4 = worksheet.Cells[i, 44].Value.ToString() ?? "",
                                    Amount_used4 = worksheet.Cells[i, 45].Value.ToString() ?? "",
                                    Unit4 = worksheet.Cells[i, 46].Value.ToString() ?? "",
                                    Material_code5 = worksheet.Cells[i, 47].Value.ToString() ?? "",
                                    Material_text5 = worksheet.Cells[i, 48].Value.ToString() ?? "",
                                    Amount_used5 = worksheet.Cells[i, 49].Value.ToString() ?? "",
                                    Unit5 = worksheet.Cells[i, 50].Value.ToString() ?? "",
                                    Material_code6 = worksheet.Cells[i, 51].Value.ToString() ?? "",
                                    Material_text6 = worksheet.Cells[i, 52].Value.ToString() ?? "",
                                    Amount_used6 = worksheet.Cells[i, 53].Value.ToString() ?? "",
                                    Unit6 = worksheet.Cells[i, 54].Value.ToString() ?? "",
                                    Material_code7 = worksheet.Cells[i, 55].Value.ToString() ?? "",
                                    Material_text7 = worksheet.Cells[i, 56].Value.ToString() ?? "",
                                    Amount_used7 = worksheet.Cells[i, 57].Value.ToString() ?? "",
                                    Unit7 = worksheet.Cells[i, 58].Value.ToString() ?? "",
                                    Material_code8 = worksheet.Cells[i, 59].Value.ToString() ?? "",
                                    Material_text8 = worksheet.Cells[i, 60].Value.ToString() ?? "",
                                    Amount_used8 = worksheet.Cells[i, 61].Value.ToString() ?? "",
                                    Unit8 = worksheet.Cells[i, 62].Value.ToString() ?? "",
                                    Inner_Code = worksheet.Cells[i, 63].Value.ToString() ?? "",
                                    Classify_Code = worksheet.Cells[i, 64].Value.ToString() ?? "",
                                });
                            }
                        }

                    }
                }

                return orders;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private List<OrderModel> SortOrders(List<OrderModel> lst, string preShape = "", int preQuantity = 0, double preDiameter = 0)
        {
            //tạo 1 list mới nhận kết quả
            var result = new List<OrderModel>();
            while (result.Count < lst.Count)
            {
                if (result.Count > 0)
                {
                    preQuantity = result.Last().TeethQuantity;
                    preDiameter = result.Last().Diameter_d;
                    preShape = result.Last().TeethShape;
                }

                preShape = Select_preShape(lst.Except(result).ToList(), preShape, preQuantity, preDiameter);
                //Nếu không có shape = preShape  ==> lấy preShape trong dòng đầu tiên trong list còn lại
                if (preShape.Equals("NotEqual"))
                    preShape = lst.Except(result).OrderBy(x => Math.Abs(preQuantity - x.TeethQuantity)).FirstOrDefault().TeethShape;
                //result.AddRange(lst.Where(x => x.shape.Equals(preShape)).OrderBy(bw => Math.Abs(preQuantity - bw.quantity)));

                var result_shape = new List<OrderModel>();
                var result_teeth = new List<OrderModel>();
                result_shape.AddRange(lst.Except(result).Where(x => x.TeethShape.Equals(preShape)).OrderBy(bw => Math.Abs(preQuantity - bw.TeethQuantity)));

                while (result_teeth.Count < result_shape.Count)
                {
                    if (result_teeth.Count > 0)
                        preDiameter = result_teeth.Last().Diameter_d;

                    preQuantity = Select_preQuantity(result_shape.Except(result_teeth).ToList(), preQuantity, preDiameter);
                    //Nếu không có số răng = preQuantity (preQuantity=-1) ==> lấy preQuantity trong dòng đầu tiên trong list còn lại
                    if (preQuantity == -1)
                        preQuantity = result_shape.Except(result_teeth).OrderBy(x => Math.Abs(preDiameter - x.Diameter_d)).FirstOrDefault().TeethQuantity;

                    result_teeth.AddRange(result_shape.Where(x => x.TeethQuantity.Equals(preQuantity)).OrderBy(bw => Math.Abs(preDiameter - bw.Diameter_d)));
                }
                result.AddRange(result_teeth);

            }

            return result;
        }

      
        public void SetZoom(System.Windows.Controls.WebBrowser WebBrowser1, double Zoom)
        {
            // For this code to work: add the Microsoft.mshtml .NET reference      
            mshtml.IHTMLDocument2 doc = WebBrowser1.Document as mshtml.IHTMLDocument2;
            doc.parentWindow.execScript("document.body.style.zoom=" + Zoom.ToString().Replace(",", ".") + ";");
        }

        //2019-11-23 Start add: add function get Teeth from TeethShape
        private string GetTeethPart(string teethShape)
        {
            string[] arr = new string[] { "S2M", "S3M", "S5M", "AT5", "MXL", "T5", "XL" };
            string teeth = "";
            foreach (var t in arr)
            {
                if (teethShape.Contains(t))
                {
                    teeth = t;
                    break;
                }
            }
            return teeth;
        }
        //2019-11-23 End add

        //2019-11-27 Start add: function Calc_PO_and_Orders_BW, Calc_PO_and_Orders_CNC
        private void Calc_PO_and_Orders_BW()
        {
            if (ListOrdersBWHob11 != null && ListOrdersBWHob11.Count >= 0)
            {
                TotalPO_BWHob11 = ListOrdersBWHob11.Count;
                TotalOrders_BWHob11 = ListOrdersBWHob11.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            if (ListOrdersBWHob12 != null && ListOrdersBWHob12.Count >= 0)
            {
                TotalPO_BWHob12 = ListOrdersBWHob12.Count;
                TotalOrders_BWHob12 = ListOrdersBWHob12.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            if (ListOrdersBWCNCHob9 != null && ListOrdersBWCNCHob9.Count >= 0)
            {
                TotalPO_BWCNCHob9 = ListOrdersBWCNCHob9.Count;
                TotalOrders_BWCNCHob9 = ListOrdersBWCNCHob9.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count >= 0)
            {
                TotalPO_BWCNCHob10 = ListOrdersBWCNCHob10.Count;
                TotalOrders_BWCNCHob10 = ListOrdersBWCNCHob10.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            if (ListOrdersBWNoHob != null && ListOrdersBWNoHob.Count >= 0)
            {
                TotalPO_BWNoHob = ListOrdersBWNoHob.Count;
                TotalOrders_BWNoHob = ListOrdersBWNoHob.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            if (ListOrdersBWF1 != null && ListOrdersBWF1.Count >= 0)
            {
                TotalPO_BWF1 = ListOrdersBWF1.Count;
                TotalOrders_BWF1 = ListOrdersBWF1.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            //2019-12-10 Start add: count total orders of BW to MySQL for checking orders changing
            var count_ordersHob11 = ListOrdersBWHob11 != null ? ListOrdersBWHob11.Count : 0;
            var count_ordersHob12 = ListOrdersBWHob12 != null ? ListOrdersBWHob12.Count : 0;
            var count_ordersNoHob = ListOrdersBWNoHob != null ? ListOrdersBWNoHob.Count : 0;
            var count_ordersF1 = ListOrdersBWF1 != null ? ListOrdersBWF1.Count : 0;
            var count_ordersBWCNCHob9 = ListOrdersBWCNCHob9 != null ? ListOrdersBWCNCHob9.Where(c => c.LineAB.Equals("BW")).ToList().Count : 0;
            var count_ordersBWCNCHob10 = ListOrdersBWCNCHob10 != null ? ListOrdersBWCNCHob10.Where(c => c.LineAB.Equals("BW")).ToList().Count : 0;
            int total = count_ordersHob11 + count_ordersHob12 + count_ordersNoHob + count_ordersF1 + count_ordersBWCNCHob9 + count_ordersBWCNCHob10;
            //try
            //{
            //    //update vào OrdersCount trong MySQL
            //    int result = OrderBLL.Instance.UpdateTotalBWOrders(total);
            //}
            //catch (Exception ex)
            //{

            //}
            //2019-12-10 End add
        }

        private void Calc_PO_and_Orders_CNC()
        {
            if (ListOrdersCNCHob9 != null && ListOrdersCNCHob9.Count >= 0)
            {
                TotalPO_CNCHob9 = ListOrdersCNCHob9.Count;
                TotalOrders_CNCHob9 = ListOrdersCNCHob9.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            if (ListOrdersCNCHob10 != null && ListOrdersCNCHob10.Count >= 0)
            {
                TotalPO_CNCHob10 = ListOrdersCNCHob10.Count;
                TotalOrders_CNCHob10 = ListOrdersCNCHob10.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            if (ListOrdersCNCNoHob != null && ListOrdersCNCNoHob.Count >= 0)
            {
                TotalPO_CNCNoHob = ListOrdersCNCNoHob.Count;
                TotalOrders_CNCNoHob = ListOrdersCNCNoHob.GroupBy(k => k.Manufa_Instruction_No).Select(g => new { key = g.Key, value = g.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions)) }).Sum(c => c.value);
            }

            //2019-12-10 Start add: count total orders of CNC to MySQL for checking orders changing
            var count_ordersHob9 = ListOrdersCNCHob9 != null ? ListOrdersCNCHob9.Count : 0;
            var count_ordersHob10 = ListOrdersCNCHob10 != null ? ListOrdersCNCHob10.Count : 0;
            var count_ordersNoHob = ListOrdersCNCNoHob != null ? ListOrdersCNCNoHob.Count : 0;
            int total = count_ordersHob9 + count_ordersHob10 + count_ordersNoHob;
            try
            {
                //update vào OrdersCount trong MySQL
                int result = OrderBLL.Instance.UpdateTotalCNCOrders(total);
            }
            catch (Exception ex)
            {
                
            }
            //2019-12-10 End add
        }
        //2019-11-27 End add

        //2019-11-28 Start add: thêm function gửi mail nếu hết đơn, function GetDiamter để tách d từ ItemName(Product)
        private void SendMailNotificationBW()
        {
          
            if (ListOrdersBWHob11 != null && ListOrdersBWHob12 != null)
            {
                var ordersBWCNCHob9 = ListOrdersBWCNCHob9 != null ? ListOrdersBWCNCHob9.Where(c => c.LineAB.Equals("BW")).ToList() : new List<OrderModel>();
                var ordersBWCNCHob10 = ListOrdersBWCNCHob10 != null ? ListOrdersBWCNCHob10.Where(c => c.LineAB.Equals("BW")).ToList() : new List<OrderModel>();

                var ordersBWHob11 = ListOrdersBWHob11 != null ? ListOrdersBWHob11.Count : 0;
                var ordersBWHob12 = ListOrdersBWHob12 != null ? ListOrdersBWHob12.Count : 0;
                //2019-12-05 Start change: nếu orders còn lại mỗi máy dưới 5 thì gửi mail PC 
                //var countOrdersBW = ListOrdersBWHob11.Count + ListOrdersBWHob12.Count + (ListOrdersBWNoHob != null ? ListOrdersBWNoHob.Count : 0);
                //var countOrdersCNC = (ordersBWCNCHob9 != null ? ordersBWCNCHob9.Count : 0) + (ordersBWCNCHob10 != null ? ordersBWCNCHob10.Count : 0);
                //if ((countOrdersBW + countOrdersCNC) < 6)
                //{
                //    IsSendingMailBW = true; //kích hoạt gửi mail tự động
                //    if (IsSendingMailBW)
                //    {
                //        string mailBody = string.Format(@"
                //        Pulley's PO is running out.\n
                //        Remaining BW's PO currently: {0}.\n
                //        Please prepare more PO.\n\n
                //        Thanks!\n
                //        (Note: This mail was sent by Pulley's PO automation tool, please do not reply!)
                //        ", (countOrdersBW + countOrdersCNC));
                //        //gửi mail báo PC
                //        int result = SPCMailHelper.Instance.PrepareSPCMail(SPCMailSender, SPCMailReceiver, "Pulley's PO automation tool", mailBody, "");
                //        if (result == 1)//gửi thành công 
                //            IsSendingMailBW = false; //tắt gửi mail tự động 
                //    }

                //}

                //gửi mail khi orders của BWHob11, BWHob12 = 4 hoặc 1
                if(ordersBWHob11 == 4 || ordersBWHob11 == 1 || ordersBWHob12 == 4 || ordersBWHob12 == 1)
                {
                    IsSendingMailBW = true; //kích hoạt gửi mail tự động
                    if (IsSendingMailBW)
                    {
                        string hob11TeethInfo = string.Format("[KnifeType: {0} | TeethQuantity: {1} | Diameter: {2}]", Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter);
                        string hob12TeethInfo = string.Format("[KnifeType: {0} | TeethQuantity: {1} | Diameter: {2}]", Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter);
                        string mailBody = string.Format(@"
                                                    Pulley's PO is running out.</br>
                                                    Remaining BW's PO currently:</br>
                                                    - Hob#11: {0}</br> 
                                                    - Teeth info: {1}</br>
                                                    - Hob#12: {2}</br> 
                                                    - Teeth info: {3}</br>
                                                    Please prepare more PO.</br></br>
                                                    Thanks!</br>
                                                    (Note: This mail was sent by Pulley's PO automation tool, please do not reply!)
                                                    ", ordersBWHob11, hob11TeethInfo, ordersBWHob12, hob12TeethInfo);
                        //gửi mail báo PC
                        
                        int result = SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "thienphuoc@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "");
                        if (result == 1)//gửi thành công 
                            IsSendingMailBW = false; //tắt gửi mail tự động 
                    }
                }

                //2019-12-05 End change

            }
        }

        private void SendMailNotificationCNC()
        {

            if (ListOrdersCNCHob9 != null && ListOrdersCNCHob10 != null)
            {
                //2019-12-05 Start change: đổi cách gửi mail PC, nếu orders của mỗi máy < 5 thì gửi mail PC
                var ordersHob9CNC = ListOrdersCNCHob9 != null ? ListOrdersCNCHob9.Count : 0;
                var ordersHob10CNC = ListOrdersCNCHob10 != null ? ListOrdersCNCHob10.Count : 0;
                //var countOrdersCNC = ListOrdersCNCHob9.Count + ListOrdersCNCHob10.Count + (ListOrdersCNCNoHob != null ? ListOrdersCNCNoHob.Count : 0);
                if (ordersHob9CNC == 4 || ordersHob9CNC == 1 || ordersHob10CNC == 4 || ordersHob10CNC == 1)
                {
                    IsSendingMailCNC = true; //kích hoạt gửi mail tự động
                    if (IsSendingMailCNC)
                    {
                        string hob9TeethInfo = string.Format("[KnifeType: {0} | TeethQuantity: {1} | Diameter: {2}]",Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                        string hob10TeethInfo = string.Format("[KnifeType: {0} | TeethQuantity: {1} | Diameter: {2}]", Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                        string mailBody = string.Format(@"
                        Pulley's PO is running out.</br>
                        Remaining CNC's PO currently:</br>
                        - Hob#9:  {0}</br> 
                        - Teeth info: {1}</br>
                        - Hob#10: {2}</br> 
                        - Teeth info: {3}</br>
                        Please prepare more PO.</br></br>
                        Thanks!</br>
                        (Note: This mail was sent by Pulley's PO automation tool, please do not reply!)
                        ", ordersHob9CNC, hob9TeethInfo,ordersHob10CNC, hob10TeethInfo);
                        //gửi mail báo PC
                        
                        int result = SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "thienphuoc@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "");
                        if (result == 1)//gửi thành công 
                            IsSendingMailCNC = false; //tắt gửi mail tự động 
                    }

                }
            }
        }

        private double GetDiameter(double d_HobShaft, string itemName)
        {
            double d = 0;
            string[] items = itemName.Split('-');
            double[] d_parts = new[] { 3, 4, 4.5, 5, 6, 6.35, 7, 8, 10, 12, 14, 15 };
            if (d_HobShaft == 3000) //tách ItemName lấy phần tử thứ 3
            {
                //get number from 3rd item
                d = GeneralUtility.Instance.SelectNumPartDouble(items[2]);
            }
            else if (d_HobShaft == 4000) //tách ItemName  lấy phần tử thứ 4
            {
                //get number from 4th item
                d = GeneralUtility.Instance.SelectNumPartDouble(items[3]);
            }
            else //lấy d_HobShaft mặc định
            {
                d = d_HobShaft;
            }
            ////lấy sai số nhỏ nhất của d với từng d_parts
            //List<double> lstSSNN = new List<double>();
            //d_parts.ToList().ForEach(c => { lstSSNN.Add(Math.Abs(c - d)); });
            //var ssnn = lstSSNN.Min();
            ////lấy d - ssnn
            //return d - ssnn;

            if (d > 15)
                d = 15;
            else if (d == 9)
                d = 8;
            else if (d == 11)
                d = 10;
            else if (d == 13)
                d = 12;

            return d;

        }

        private string Select_preShape(List<OrderModel> lst, string preShape = "", int preQuantity = 0, double preDiameter = 0)
        {
            //Nếu không có shape = preShape 
            if (lst.Where(x => x.TeethShape.Equals(preShape)).Count() == 0)
            {
                //==> Đổi preShape theo item có số răng = preQuantity (nếu có)
                if (lst.Where(x => x.TeethQuantity.Equals(preQuantity)).Count() > 0)
                {
                    preShape = lst.Where(x => x.TeethQuantity.Equals(preQuantity)).FirstOrDefault().TeethShape;
                }
                else
                {
                    //==> Đổi preShape theo item có đường kính = preDiameter (nếu có)
                    if (lst.Where(x => x.Diameter_d.Equals(preDiameter)).Count() > 0)
                    {
                        preShape = lst.Where(x => x.Diameter_d.Equals(preDiameter)).FirstOrDefault().TeethShape;
                    }
                    else
                    {
                        preShape = "NotEqual";
                    }
                }
            }

            return preShape;
        }

        private int Select_preQuantity(List<OrderModel> lst, int preQuantity = 0, double preDiameter = 0)
        {
            //Nếu không có số răng = preQuantity 
            if (lst.Where(x => x.TeethQuantity.Equals(preQuantity)).Count() == 0)
            {
                //==> Đổi preQuantity theo item có đường kính = preDiameter (nếu có)
                if (lst.Where(x => x.Diameter_d.Equals(preDiameter)).Count() > 0)
                {
                    preQuantity = lst.Where(x => x.Diameter_d.Equals(preDiameter)).FirstOrDefault().TeethQuantity;
                }
                else
                {
                    preQuantity = -1;
                }
            }

            return preQuantity;
        }

        private string Select_preDMat(List<OrderModel> lst, string preShape = "", int preQuantity = 0, double preDiameter = 0, string preDMat = "")
        {

            //Nếu không có DMat = preDMat 
            if (lst.Where(x => x.Material_text1.Equals(preDMat)).Count() == 0)
            {
                //==> Đổi preDMat theo item có TeethShape = preShape (nếu có)
                if (lst.Where(x => x.TeethShape.Equals(preShape)).Count() > 0)
                {
                    //preShape = lst.Where(x => x.TeethQuantity.Equals(preQuantity)).FirstOrDefault().TeethShape;
                    preDMat = lst.Where(x=>x.TeethShape.Equals(preShape)).FirstOrDefault().Material_text1;
                }
                else
                {
                    //==> Đổi preDMat theo item có TeethQuantity = preQuantity (nếu có)
                    if (lst.Where(x => x.TeethQuantity == preQuantity).Count() > 0)
                    {
                        //preShape = lst.Where(x => x.Diameter_d.Equals(preDiameter)).FirstOrDefault().TeethShape;
                        preDMat = lst.Where(x => x.TeethQuantity == preQuantity).FirstOrDefault().Material_text1;
                    }
                    else
                    {
                        //==> Đổi preDMat theo item có Diameter_d = preDiameter (nếu có)
                        if(lst.Where(x=>x.Diameter_d == preDiameter).Count() > 0)
                        {
                            preDMat = lst.Where(x => x.Diameter_d == preDiameter).FirstOrDefault().Material_text1;
                        }
                        else
                        {
                            preDMat = "NotEqual";
                        }

                    }
                }
            }

            return preDMat;
        }

        //2020-01-04 Start add: functions sắp xếp theo đường kính vật liệu cho máy CNC
        private static List<OrderModel> SortOrders_CNC(List<OrderModel> lst_initial, string preMaterial = "", string preShape = "", int preQuantity = 0, double preDiameter = 0)
        {
            List<OrderModel> result_final = new List<OrderModel>();
            while (result_final.Count < lst_initial.Count)
            {
                if (result_final.Count > 0)
                {
                    preMaterial = result_final.Last().Material_text1;
                    preShape = result_final.Last().TeethShape;
                    preQuantity = result_final.Last().TeethQuantity;
                    preDiameter = result_final.Last().Diameter_d;
                }
                preMaterial = Select_preMaterial_CNC(lst_initial.Except(result_final).ToList(), preMaterial, preShape, preQuantity, preDiameter);
                //Nếu không có shape = preShape  ==> lấy preShape trong dòng đầu tiên trong list còn lại
                //2020-01-04: hàm Select_preQuantity đã xử lý rồi
                //if (preMaterial == "NotEqual")
                //    preMaterial = lst_initial.Except(result_final).OrderBy(x => x.material).FirstOrDefault().material;
                //Danh sách đã lọc theo D vật liệu

                var lst = new List<OrderModel>();
                lst.AddRange(lst_initial.Except(result_final).Where(x => x.Material_text1.Equals(preMaterial)));

                var result = new List<OrderModel>();

                while (result.Count < lst.Count)
                {
                    if (result.Count > 0)
                    {
                        preShape = result.Last().TeethShape;
                        preQuantity = result.Last().TeethQuantity;
                        preDiameter = result.Last().Diameter_d;

                    }

                    preShape = Select_preShape_CNC(lst.Except(result).ToList(), preShape, preQuantity, preDiameter);
                    //Nếu không có shape = preShape  ==> lấy preShape trong dòng đầu tiên trong list còn lại
                    //2020-01-04: hàm Select_preQuantity đã xử lý rồi
                    //if (preShape == "NotEqual")
                    //    preShape = lst.Except(result).OrderBy(x => Math.Abs(preQuantity - x.quantity)).FirstOrDefault().shape;
                    //result.AddRange(lst.Where(x => x.shape.Equals(preShape)).OrderBy(OrderModel => Math.Abs(preQuantity - OrderModel.quantity)));

                    var result_shape = new List<OrderModel>();
                    var result_teeth = new List<OrderModel>();
                    result_shape.AddRange(lst.Except(result).Where(x => x.TeethShape.Equals(preShape)).OrderBy(OrderModel => Math.Abs(preQuantity - OrderModel.TeethQuantity)));

                    while (result_teeth.Count < result_shape.Count)
                    {
                        if (result_teeth.Count > 0)
                            preDiameter = result_teeth.Last().Diameter_d;

                        preQuantity = Select_preQuantity_CNC(result_shape.Except(result_teeth).ToList(), preQuantity, preDiameter);
                        //Nếu không có số răng = preQuantity (preQuantity=-1) ==> lấy preQuantity trong dòng đầu tiên trong list còn lại
                        //2020-01-04: hàm Select_preQuantity đã xử lý rồi
                        //if (preQuantity == -1)
                        //    preQuantity = result_shape.Except(result_teeth).OrderBy(x => Math.Abs(preDiameter - x.diameter)).FirstOrDefault().quantity;

                        result_teeth.AddRange(result_shape.Where(x => x.TeethQuantity.Equals(preQuantity)).OrderBy(OrderModel => Math.Abs(preDiameter - OrderModel.Diameter_d)));
                    }
                    result.AddRange(result_teeth);

                }
                result_final.AddRange(result);

            }
            return result_final;
        }
        
        /// <summary>
        /// sort order SSD (N1/N) for CNC machines
        /// </summary>
        /// <param name="orders"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private List<OrderModel> SetupSSD_CNC(List<OrderModel> orders, SSDType ssdType, string material = "", string shape = "", int qty = 0, double d = 0)
        {
            int dayToAddN1 = HolidayHelper.Instance.DayCountN1;
            List<OrderModel> ordersToSave = new List<OrderModel>();
            if (ssdType == SSDType.N1)
            {
                //22-10-2020: cập nhật: check có hàng N thì ưu tiên trước N1
                var ordersN = orders.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                var ordersN1 = orders.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();
                //có đơn N
                if(ordersN != null && ordersN.Count > 0)
                {
                    var sorted_ordersN = SortOrders_CNC(ordersN, material, shape, qty, d);
                    var last_orderN = sorted_ordersN.LastOrDefault();
                    ordersToSave.AddRange(sorted_ordersN);
                    if(ordersN1 != null && ordersN1.Count > 0)//có đơn N1
                    {
                        var sorted_ordersN1 = SortOrders_CNC(ordersN1, last_orderN.Material_text1, last_orderN.TeethShape, last_orderN.TeethQuantity, last_orderN.Diameter_d);
                        ordersToSave.AddRange(sorted_ordersN1);
                        var last_sorted_orderN1 = sorted_ordersN1.LastOrDefault();
                        var remain_orders = orders.Except(ordersN).Except(ordersN1).ToList();
                        if(remain_orders != null && remain_orders.Count > 0)
                        {
                            var sorted_remain_orders = SortOrders_CNC(remain_orders, last_sorted_orderN1.Material_text1, last_sorted_orderN1.TeethShape, last_sorted_orderN1.TeethQuantity, last_sorted_orderN1.Diameter_d);
                            ordersToSave.AddRange(sorted_remain_orders);
                        }
                    }
                    else
                    {
                        //chỉ còn đơn normal
                        var remain_orders = orders.Except(ordersN).ToList();
                        if(remain_orders != null && remain_orders.Count > 0)
                        {
                            var sorted_remain_orders = SortOrders_CNC(remain_orders, last_orderN.Material_text1, last_orderN.TeethShape, last_orderN.TeethQuantity, last_orderN.Diameter_d);
                            ordersToSave.AddRange(sorted_remain_orders);
                        }
                    }
                }
                else if(ordersN1 != null && ordersN1.Count > 0)
                {
                    //có N1 CNC
                    //var groupOrdersN1_CNC = orders.GroupBy(g => g.TeethShape).Select(g => new { gKey = g.Key, gItems = g.ToList().Where(c => c.LineAB.Equals("CNC")).ToList() }).ToList();
                    var sorted_ordersN1 = SortOrders_CNC(ordersN1, material, shape, qty, d);
                    ordersToSave.AddRange(sorted_ordersN1);
                    var last_sorted_ordersN1 = sorted_ordersN1.LastOrDefault();
                    var remain_orders = orders.Except(ordersN1).ToList();
                    if(remain_orders != null && remain_orders.Count > 0)
                    {
                        var sorted_remain_orders = SortOrders_CNC(remain_orders, last_sorted_ordersN1.Material_text1, last_sorted_ordersN1.TeethShape, last_sorted_ordersN1.TeethQuantity, last_sorted_ordersN1.Diameter_d);
                        ordersToSave.AddRange(sorted_remain_orders);
                    }

                }
                else
                {
                    //ko có N1 CNC thì xếp thường
                    var sorted_orders = SortOrders_CNC(orders, material, shape, qty, d);
                    ordersToSave.AddRange(sorted_orders);
                }
                
            }
            return ordersToSave;
        }

        private List<OrderModel> SetupSSD_BW(List<OrderModel> orders, SSDType ssdType, string shape, int teethQty, double d)
        {
            int dayToAddN1 = HolidayHelper.Instance.DayCountN1;
            List<OrderModel> ordersToSave = new List<OrderModel>();
            if (ssdType == SSDType.N1)
            {
                //22-10-2020: cập nhật: check có hàng N thì ưu tiên trước rồi tới N+1
                var ordersN = orders.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                var ordersN1 = orders.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();

                //có orders N
                if(ordersN != null && ordersN.Count > 0)
                {
                    var sorted_ordersN = SortOrders(ordersN, shape, teethQty, d);
                    var last_orderN = sorted_ordersN.LastOrDefault();
                    ordersToSave.AddRange(sorted_ordersN);
                    if(ordersN1 != null && ordersN1.Count > 0)//có orders N1
                    {
                        var sorted_ordersN1 = SortOrders(ordersN1, last_orderN.TeethShape, last_orderN.TeethQuantity, last_orderN.Diameter_d);
                        ordersToSave.AddRange(sorted_ordersN1);
                        var last_orderN1 = sorted_ordersN1.LastOrDefault();
                        var remain_orders = orders.Except(ordersN).Except(ordersN1).ToList();
                        if(remain_orders != null && remain_orders.Count > 0)
                        {
                            var sorted_remain_orders = SortOrders(remain_orders, last_orderN1.TeethShape, last_orderN1.TeethQuantity, last_orderN1.Diameter_d);
                            ordersToSave.AddRange(sorted_remain_orders);
                        }
                    }
                    else
                    {
                        //chỉ còn orders normal
                        var remain_orders = orders.Except(ordersN).ToList();
                        if(remain_orders != null && remain_orders.Count > 0)
                        {
                            var sorted_remain_orders = SortOrders(remain_orders, last_orderN.TeethShape, last_orderN.TeethQuantity, last_orderN.Diameter_d);
                            ordersToSave.AddRange(sorted_remain_orders);
                        }
                    }
                }
                else if (ordersN1 != null && ordersN1.Count > 0)
                {
                    //có N1 BW
                    //var groupOrdersN1_CNC = orders.GroupBy(g => g.TeethShape).Select(g => new { gKey = g.Key, gItems = g.ToList().Where(c => c.LineAB.Equals("CNC")).ToList() }).ToList();
                    var sorted_ordersN1 = SortOrders(ordersN1, shape, teethQty, d);
                    ordersToSave.AddRange(sorted_ordersN1);
                    var last_sorted_ordersN1 = sorted_ordersN1.LastOrDefault();
                    var remain_orders = orders.Except(ordersN1).ToList();
                    if (remain_orders != null && remain_orders.Count > 0)
                    {
                        var sorted_remain_orders = SortOrders(remain_orders,  last_sorted_ordersN1.TeethShape, last_sorted_ordersN1.TeethQuantity, last_sorted_ordersN1.Diameter_d);
                        ordersToSave.AddRange(sorted_remain_orders);
                    }

                }
                else
                {
                    //ko có N và N1 BW thì xếp thường
                    var sorted_orders = SortOrders(orders, shape, teethQty, d);
                    ordersToSave.AddRange(sorted_orders);
                }

            }
            return ordersToSave;
        }

        static string Select_preMaterial_CNC(List<OrderModel> lst, string preMaterial = "", string preShape = "", int preQuantity = 0, double preDiameter = 0)
        {
            //Nếu không có shape = preShape 
            if (lst.Where(x => x.Material_text1.Equals(preMaterial)).Count() == 0)
            {
                //==> Đổi preMaterial theo item có shape = preShape (nếu có)
                if (lst.Where(x => x.TeethShape.Equals(preShape)).Count() > 0)
                {
                    preMaterial = lst.Where(x => x.TeethShape.Equals(preShape)).FirstOrDefault().Material_text1;
                }
                //Nếu không có shape = preShape
                else
                {
                    //==> Đổi preMaterial theo item có số răng = preQuantity (nếu có)
                    if (lst.Where(x => x.TeethQuantity.Equals(preQuantity)).Count() > 0)
                    {
                        preMaterial = lst.Where(x => x.TeethQuantity.Equals(preQuantity)).FirstOrDefault().Material_text1;
                    }
                    else
                    {
                        //==> Đổi preMaterial theo item có đường kính = preDiameter (nếu có)
                        if (lst.Where(x => x.Diameter_d.Equals(preDiameter)).Count() > 0)
                        {
                            preMaterial = lst.Where(x => x.Diameter_d.Equals(preDiameter)).FirstOrDefault().Material_text1;
                        }
                        else
                        {
                            //preMaterial = "NotEqual";
                            preMaterial = lst.OrderBy(x => x.Material_text1).ThenBy(c=>c.Factory_Ship_Date).FirstOrDefault().Material_text1;
                        }
                    }
                }
            }
            return preMaterial;
        }
        static string Select_preShape_CNC(List<OrderModel> lst, string preShape = "", int preQuantity = 0, double preDiameter = 0)
        {
            //Nếu không có shape = preShape 
            if (lst.Where(x => x.TeethShape.Equals(preShape)).Count() == 0)
            {
                //==> Đổi preShape theo item có số răng = preQuantity (nếu có)
                if (lst.Where(x => x.TeethQuantity.Equals(preQuantity)).Count() > 0)
                {
                    preShape = lst.Where(x => x.TeethQuantity.Equals(preQuantity)).FirstOrDefault().TeethShape;
                }
                else
                {
                    //==> Đổi preShape theo item có đường kính = preDiameter (nếu có)
                    if (lst.Where(x => x.Diameter_d.Equals(preDiameter)).Count() > 0)
                    {
                        preShape = lst.Where(x => x.Diameter_d.Equals(preDiameter)).FirstOrDefault().TeethShape;
                    }
                    else
                    {
                        //preShape = "NotEqual";
                        preShape = lst.OrderBy(x => Math.Abs(preQuantity - x.TeethQuantity)).ThenBy(c => c.Factory_Ship_Date).FirstOrDefault().TeethShape;
                    }
                }
            }
            return preShape;
        }
        static int Select_preQuantity_CNC(List<OrderModel> lst, int preQuantity = 0, double preDiameter = 0)
        {
            //Nếu không có số răng = preQuantity 
            if (lst.Where(x => x.TeethQuantity.Equals(preQuantity)).Count() == 0)
            {
                ////==> Đổi preQuantity theo item có đường kính = preDiameter (nếu có)
                /////Nếu đường kính = preDiameter ==> preDiameter - x.diameter =0 ==> Khỏi cần dùng if else 2020-01-04
                preQuantity = lst.OrderBy(x => Math.Abs(preDiameter - x.Diameter_d)).ThenBy(c => c.Factory_Ship_Date).FirstOrDefault().TeethQuantity;
                //if (lst.Where(x => x.diameter.Equals(preDiameter)).Count() > 0)
                //{
                //    preQuantity = lst.Where(x => x.diameter.Equals(preDiameter)).FirstOrDefault().quantity;
                //}
                //else
                //{
                //    //preQuantity = -1;
                //    preQuantity = lst.OrderBy(x => Math.Abs(preDiameter - x.diameter)).FirstOrDefault().quantity;
                //}
            }
            return preQuantity;
        }
        //2020-01-04 End add

        private string GetKnifeType(string teethShape, int teethQty)
        {
            string knifeType = "";
            try
            {
                knifeType = OrderBLL.Instance.GetKnifeType(teethShape, teethQty);
            }
            catch (Exception ex)
            {

                MessageBox.Show("Điều kiện tìm loại dao chưa đúng hoặc lỗi đường truyền. Vui lòng thử lại!\n"+ex.Message);
            }
            return knifeType;
        }
        //2019-11-28 End add

        //2019-12-16 Start add: 
        //function sắp xếp ưu tiên orders BW N, N+1 
        private void SetupOrdersBW_N1(List<OrderModel> lstOrdersN1,string orderType)
        {
                      
            StringBuilder mailBody = new StringBuilder();
            int status = 0;

            //lấy tất cả đơn N+1 rồi chia đều cho 4 máy
            //group tất cả các đơn N1 theo TeethShape và chia ra N+1 của BW, N+1 của CNC
            //var groupOrdersN1 = lstOrdersN1.GroupBy(g => g.TeethShape).Select(g => new { gKey = g.Key, gItems = g.ToList() }).ToList();
            var groupOrdersN1_BW = lstOrdersN1.GroupBy(g => g.TeethShape).Select(g => new { gKey = g.Key, gItems = g.ToList().Where(c => c.LineAB.Equals("BW")).ToList() }).ToList();
            var groupOrdersN1_CNC = lstOrdersN1.GroupBy(g => g.TeethShape).Select(g => new { gKey = g.Key, gItems = g.ToList().Where(c => c.LineAB.Equals("CNC")).ToList() }).ToList();

            List<OrderModel> tempList = new List<OrderModel>();
            //chia mỗi máy 1 group 
            //-----------Đầu tiên chia hết các đơn N+1 của BW cho cả máy CNC
            //orders N+1 for BW11
            ListOrdersBWHob11 = ListOrdersBWHob11 != null ? ListOrdersBWHob11 : new ObservableCollection<OrderModel>();
            var tempOdersBW11 = groupOrdersN1_BW.Where(g => g.gKey.Equals(Hob11BW_TeethShape)).FirstOrDefault();
            if (tempOdersBW11 != null && groupOrdersN1_BW.Count > 0)
            {
                var odN1bw11 = SortOrders(tempOdersBW11.gItems.ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter).ToList();
                if(odN1bw11 != null && odN1bw11.Count > 0)
                {
                    var lastOrder_BW11 = odN1bw11.LastOrDefault();
                    var remainListBW11 = ListOrdersBWHob11.Except(odN1bw11).ToList();
                    var sortedOrders_BW11 = SortOrders(remainListBW11, lastOrder_BW11.TeethShape, lastOrder_BW11.TeethQuantity, lastOrder_BW11.Diameter_d);
                    tempList.AddRange(odN1bw11);
                    tempList.AddRange(sortedOrders_BW11);
                    //remove duplicate with same PO
                    tempList.GroupBy(x => x.Manufa_Instruction_No).Select(x => x.First());
                    //refresh BW11
                    ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(tempList, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));

                    //remove group order đã chuyển cho BW11
                    groupOrdersN1_BW.Remove(tempOdersBW11);
                }
              
            }

            //orders N+1 for BW12
            ListOrdersBWHob12 = ListOrdersBWHob12 != null ? ListOrdersBWHob12 : new ObservableCollection<OrderModel>();
            var tempOdersBW12 = groupOrdersN1_BW.Where(g => g.gKey.Equals(Hob12BW_TeethShape)).FirstOrDefault();
            if (tempOdersBW12 != null && groupOrdersN1_BW.Count > 0)
            {
                var odN1bw12 = SortOrders(tempOdersBW12.gItems.ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter).ToList();
                if(odN1bw12!=null && odN1bw12.Count > 0)
                {
                    var lastOrder_BW12 = odN1bw12.LastOrDefault();
                    var remainListBW12 = ListOrdersBWHob12.Except(odN1bw12).ToList();
                    var sortedOrders_BW12 = SortOrders(remainListBW12, lastOrder_BW12.TeethShape, lastOrder_BW12.TeethQuantity, lastOrder_BW12.Diameter_d);
                    tempList = new List<OrderModel>();
                    tempList.AddRange(odN1bw12);
                    tempList.AddRange(sortedOrders_BW12);
                    tempList.GroupBy(x => x.Manufa_Instruction_No).Select(x => x.First());
                    //refresh BW12
                    ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(tempList, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));

                    //remove group order đã chuyển cho BW12
                    groupOrdersN1_BW.Remove(tempOdersBW12);
                }
             
            }
            //orders N+1 for BWCNC9
            ListOrdersBWCNCHob9 = ListOrdersBWCNCHob9 != null ? ListOrdersBWCNCHob9 : new ObservableCollection<OrderModel>();
            var tempOdersBWCNC9 = groupOrdersN1_BW.Where(g => g.gKey.Equals(Hob9BWCNC_TeethShape)).FirstOrDefault();
            if (tempOdersBWCNC9 != null && groupOrdersN1_BW.Count>0)
            {
                var odN1bwcmc9 = SortOrders(tempOdersBWCNC9.gItems.ToList(), Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter).ToList();
                if(odN1bwcmc9!=null && odN1bwcmc9.Count > 0)
                {
                    var lastOrder_BWCNC9 = odN1bwcmc9.LastOrDefault();
                    var remainListBWCNC9 = ListOrdersBWCNCHob9.Except(odN1bwcmc9).ToList();
                    var sortedOrders_BWCNC9 = SortOrders(remainListBWCNC9, lastOrder_BWCNC9.TeethShape, lastOrder_BWCNC9.TeethQuantity, lastOrder_BWCNC9.Diameter_d);
                    tempList = new List<OrderModel>();
                    tempList.AddRange(odN1bwcmc9);
                    //loại các đơn trùng từ BWHob11 và BWHob12
                    tempList.AddRange(sortedOrders_BWCNC9.Except(ListOrdersBWHob11.ToList()).Except(ListOrdersBWHob12.ToList()));
                    tempList.GroupBy(x => x.Manufa_Instruction_No).Select(x => x.First());
                    //refresh BWCNC9
                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(tempList, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));

                    //cập nhật khóa nút PRINT bên màn hình CNC

                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);

                    //remove group order đã chuyển cho BWCNC9
                    groupOrdersN1_BW.Remove(tempOdersBWCNC9);
                }
             
            }
            //orders N+1 for BWCNC10
            ListOrdersBWCNCHob10 = ListOrdersBWCNCHob10 != null ? ListOrdersBWCNCHob10 : new ObservableCollection<OrderModel>();
            var tempOdersBWCNC10 = groupOrdersN1_BW.Where(g => g.gKey.Equals(Hob10BWCNC_TeethShape)).FirstOrDefault();
            if (tempOdersBWCNC10 != null && groupOrdersN1_BW.Count>0)
            {
                var odN1bwcmc10 = SortOrders(tempOdersBWCNC10.gItems.ToList(), Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter).ToList();
                if(odN1bwcmc10 != null && odN1bwcmc10.Count > 0)
                {
                    var lastOrder_BWCNC10 = odN1bwcmc10.LastOrDefault();
                    var remainListBWCNC10 = ListOrdersBWCNCHob10.Except(odN1bwcmc10).ToList();
                    var sortedOrders_BWCNC10 = SortOrders(remainListBWCNC10, lastOrder_BWCNC10.TeethShape, lastOrder_BWCNC10.TeethQuantity, lastOrder_BWCNC10.Diameter_d);
                    tempList = new List<OrderModel>();
                    tempList.AddRange(odN1bwcmc10);
                    //loại các đơn trùng từ BWHob11 và BWHob12
                    tempList.AddRange(sortedOrders_BWCNC10.Except(ListOrdersBWHob11.ToList()).Except(ListOrdersBWHob12.ToList()));
                    tempList.GroupBy(x => x.Manufa_Instruction_No).Select(x => x.First());
                    //refresh BWCNC9
                    ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(tempList, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));

                    //cập nhật khóa nút PRINT bên màn hình CNC
                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);

                    //remove group order đã chuyển cho BWCNC10
                    groupOrdersN1_BW.Remove(tempOdersBWCNC10);
                }
            
            }

            //đếm các groupOrdersN1 còn dư
            //var countGroupsN1 = groupOrdersN1.Count;
            var countGroupsN1_BW = groupOrdersN1_BW.Count;
            
            //------------Nếu còn groupOrdersBW--------------------
            if (countGroupsN1_BW > 0)
            {               
                //chia đều group ra 
                //nếu máy BW11 ko có N+1
                if (tempOdersBW11 == null)
                {
                    //chuyển cho BW11 1 group
                    var ordersN1 = groupOrdersN1_BW.FirstOrDefault().gItems.ToList();
                    var sortedN1Orders = SortOrders(ordersN1, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter);
                    if(sortedN1Orders != null && sortedN1Orders.Count > 0)
                    {
                        var lastN1Orders = sortedN1Orders.LastOrDefault();
                        var currentOrdersBW11 = ListOrdersBWHob11.ToList();
                        var sortedNormalOrders = SortOrders(currentOrdersBW11, lastN1Orders.TeethShape, lastN1Orders.TeethQuantity, lastN1Orders.Diameter_d);
                        var res = new List<OrderModel>();
                        res.AddRange(sortedN1Orders);
                        res.AddRange(sortedNormalOrders);
                        res.GroupBy(x => x.Manufa_Instruction_No).Select(x => x.First());
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(res,Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                        //remove group orders đã xếp
                        groupOrdersN1_BW.Remove(groupOrdersN1_BW.FirstOrDefault());
                    }
                 
                }

                //nếu máy BW12 ko có N+1
                if (tempOdersBW12 == null && groupOrdersN1_BW.Count > 0)
                {
                    //chuyển cho BW12 1 group
                    var ordersN1 = groupOrdersN1_BW.FirstOrDefault().gItems.ToList();
                    var sortedN1Orders = SortOrders(ordersN1, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter);
                    if(sortedN1Orders != null && sortedN1Orders.Count > 0)
                    {
                        var lastN1Orders = sortedN1Orders.LastOrDefault();
                        var currentOrdersBW12 = ListOrdersBWHob12.ToList();
                        var sortedNormalOrders = SortOrders(currentOrdersBW12, lastN1Orders.TeethShape, lastN1Orders.TeethQuantity, lastN1Orders.Diameter_d);
                        var res = new List<OrderModel>();
                        res.AddRange(sortedN1Orders);
                        res.AddRange(sortedNormalOrders);
                        res.GroupBy(x => x.Manufa_Instruction_No).Select(x => x.First());
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(res, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                        //remove group orders đã xếp
                        groupOrdersN1_BW.Remove(groupOrdersN1_BW.FirstOrDefault());
                    }
                
                }

               

                if(groupOrdersN1_BW.Count > 0)//vẫn còn group orders N1
                {
                    //gán vào cuối list của BW11
                    var lastOrderBW11 = ListOrdersBWHob11.LastOrDefault();
                    var listOrders = new List<OrderModel>();
                    groupOrdersN1_BW.ForEach(g => listOrders.AddRange(g.gItems));
                    if(lastOrderBW11 !=null)//nếu có item thì xếp theo teethInfo
                    {
                        var sortedN1Orders = SortOrders(listOrders,lastOrderBW11.TeethShape, lastOrderBW11.TeethQuantity, lastOrderBW11.Diameter_d);
                        var currentOrdersBW11 = ListOrdersBWHob11.ToList();
                        currentOrdersBW11.AddRange(sortedN1Orders);
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(currentOrdersBW11, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                    }
                    else
                    {
                        //BW11 hết orders => add hết vào BW11
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(listOrders, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                    }
                }
            }

            //---------------Xếp các orders N+1 CNC-----------------------
            //tạm thời xếp cuối list BWCNCHob9 và BWCNCHob10

            if (groupOrdersN1_CNC.Count > 0)
            {
                //lấy last order của BWCNCHob9
                var lastOrderBWCNCHob9 = ListOrdersBWCNCHob9.LastOrDefault();
                if(lastOrderBWCNCHob9 != null)
                {
                    var tempOrdersForBWCNC9 = groupOrdersN1_CNC.Where(c => c.gKey.Equals(lastOrderBWCNCHob9.TeethShape)).FirstOrDefault();
                    if(tempOrdersForBWCNC9 != null && groupOrdersN1_CNC.Count > 0)
                    {
                        var currentOrdersBWCNCHob9 = ListOrdersBWCNCHob9.ToList();
                        var sortedTempOrdersBWCNCHob9 = SortOrders(tempOrdersForBWCNC9.gItems.ToList(), lastOrderBWCNCHob9.TeethShape, lastOrderBWCNCHob9.TeethQuantity, lastOrderBWCNCHob9.Diameter_d);
                        currentOrdersBWCNCHob9.AddRange(sortedTempOrdersBWCNCHob9);
                        //refresh ListOrdersBWCNCHob9
                        ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(currentOrdersBWCNCHob9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));

                        //remove group đã chuyển
                        groupOrdersN1_CNC.Remove(tempOrdersForBWCNC9);
                    }
                }
                else
                {
                    //nếu ko có last order tức là List trống
                    //vậy lấy orders giống TeethShape hiện tại của BWCNC9
                    var tempOrdersForBWCNCHob9 = groupOrdersN1_CNC.Where(c => c.gKey.Equals(Hob9BWCNC_TeethShape)).FirstOrDefault();
                    if(tempOrdersForBWCNCHob9 != null && groupOrdersN1_CNC.Count > 0)
                    {
                        //refresh ListOrdersBWCNCHob9
                        ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(tempOrdersForBWCNCHob9.gItems.ToList(), Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter), Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));

                        //remove group đã chuyển
                        groupOrdersN1_CNC.Remove(tempOrdersForBWCNCHob9);
                    }
                    else
                    {
                        //nếu ko cùng TeethShape thì xếp hết luôn cho BWCNCHob9
                        List<OrderModel> orders = new List<OrderModel>();
                        groupOrdersN1_CNC.ForEach(c => orders.AddRange(c.gItems.ToList()));
                        var sortedOrders = SortOrders(orders, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                        //refresh ListOrdersBWCNCHob9
                        ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(sortedOrders, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));
                        //remove groupOrdersN1_CNC
                        groupOrdersN1_CNC.Clear();
                    }
                }


                //lấy last order của BWCNCHob10
                var lastOrderBWCNCHob10 = ListOrdersBWCNCHob10.LastOrDefault();
                if (lastOrderBWCNCHob10 != null)
                {
                    var tempOrdersForBWCNC10 = groupOrdersN1_CNC.Where(c => c.gKey.Equals(lastOrderBWCNCHob10.TeethShape)).FirstOrDefault();
                    if (tempOrdersForBWCNC10 != null && groupOrdersN1_CNC.Count > 0)
                    {
                        var currentOrdersBWCNCHob10 = ListOrdersBWCNCHob10.ToList();
                        var sortedTempOrdersBWCNCHob10 = SortOrders(tempOrdersForBWCNC10.gItems.ToList(), lastOrderBWCNCHob10.TeethShape, lastOrderBWCNCHob10.TeethQuantity, lastOrderBWCNCHob10.Diameter_d);
                        currentOrdersBWCNCHob10.AddRange(sortedTempOrdersBWCNCHob10);
                        //refresh ListOrdersBWCNCHob10
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(currentOrdersBWCNCHob10, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));

                        //remove group đã chuyển
                        groupOrdersN1_CNC.Remove(tempOrdersForBWCNC10);
                    }
                }
                else
                {
                    //nếu ko có last order tức là List trống
                    //vậy lấy orders giống TeethShape hiện tại của BWCNC10
                    var tempOrdersForBWCNCHob10 = groupOrdersN1_CNC.Where(c => c.gKey.Equals(Hob10BWCNC_TeethShape)).FirstOrDefault();
                    if (tempOrdersForBWCNCHob10 != null && groupOrdersN1_CNC.Count > 0)
                    {
                        //refresh ListOrdersBWCNCHob10
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(tempOrdersForBWCNCHob10.gItems.ToList(), Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter), Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));

                        //remove group đã chuyển
                        groupOrdersN1_CNC.Remove(tempOrdersForBWCNCHob10);
                    }
                    else
                    {
                        //nếu ko cùng TeethShape thì xếp hết luôn cho BWCNCHob10
                        List<OrderModel> orders = new List<OrderModel>();
                        groupOrdersN1_CNC.ForEach(c => orders.AddRange(c.gItems.ToList()));
                        var sortedOrders = SortOrders(orders, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        //refresh ListOrdersBWCNCHob9
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(sortedOrders, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //remove groupOrdersN1_CNC
                        groupOrdersN1_CNC.Clear();
                    }
                }


            }


            var countGroupsN1_CNC = groupOrdersN1_CNC.Count;
            //nếu groupOrdersN1_CNC vẫn còn
            if(countGroupsN1_CNC > 0)
            {
                //chia đơn cho BWCNC9
                var lastOrderBWCNCHob9 = ListOrdersBWCNCHob9.LastOrDefault();
               
                //nếu có last order của BWCNC9
                if(lastOrderBWCNCHob9 != null)
                {
                    var ordersFromGroupCNC = groupOrdersN1_CNC.Where(c => c.gKey.Equals(lastOrderBWCNCHob9.TeethShape)).ToList();
                    List<OrderModel> orders = new List<OrderModel>();
                    ordersFromGroupCNC.ForEach(c => orders.AddRange(c.gItems.ToList()));
                    var sortedOrders = SortOrders(orders, lastOrderBWCNCHob9.TeethShape, lastOrderBWCNCHob9.TeethQuantity, lastOrderBWCNCHob9.Diameter_d);
                    var currentListBWCNCHob9 = ListOrdersBWCNCHob9.ToList();
                    currentListBWCNCHob9.AddRange(sortedOrders);
                    //refresh ListOrdersBWCNCHob9
                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(currentListBWCNCHob9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));
                    //remove group đã chia cho BWCNCHob9
                    groupOrdersN1_CNC = groupOrdersN1_CNC.Except(ordersFromGroupCNC).ToList();
                }
                else
                {
                    //nếu ko có last order thì lấy ra các group có order cùng TeethShape
                    var ordersFromGroupCNC = groupOrdersN1_CNC.Where(c => c.gKey.Equals(lastOrderBWCNCHob9.TeethShape)).ToList();
                    List<OrderModel> orders = new List<OrderModel>();
                    ordersFromGroupCNC.ForEach(c => orders.AddRange(c.gItems.ToList()));                    
                    var sortedOrders = SortOrders(orders, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                    //refresh ListOrdersBWCNCHob9
                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(sortedOrders, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));
                    //remove group đã chia cho BWCNCHob9
                    groupOrdersN1_CNC = groupOrdersN1_CNC.Except(ordersFromGroupCNC).ToList();
                }

                //còn lại chia hết cho BWCNCHob10
                if(groupOrdersN1_CNC.Count > 0)
                {
                    var lastOrderBWCNCHob10 = ListOrdersBWCNCHob10.LastOrDefault();
                    //nếu có last order của BWCNC10
                    if (lastOrderBWCNCHob10 != null)
                    {
                        var ordersFromGroupCNC = groupOrdersN1_CNC.Where(c => c.gKey.Equals(lastOrderBWCNCHob10.TeethShape)).ToList();
                        List<OrderModel> orders = new List<OrderModel>();
                        ordersFromGroupCNC.ForEach(c => orders.AddRange(c.gItems.ToList()));
                        var sortedOrders = SortOrders(orders, lastOrderBWCNCHob10.TeethShape, lastOrderBWCNCHob10.TeethQuantity, lastOrderBWCNCHob10.Diameter_d);
                        var currentListBWCNCHob10 = ListOrdersBWCNCHob10.ToList();
                        currentListBWCNCHob10.AddRange(sortedOrders);
                        //refresh ListOrdersBWCNCHob10
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(currentListBWCNCHob10, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //remove group đã chia cho BWCNCHob10
                        groupOrdersN1_CNC = groupOrdersN1_CNC.Except(ordersFromGroupCNC).ToList();
                    }
                    else
                    {
                        //nếu ko có last order thì lấy ra các group có order cùng TeethShape
                        var ordersFromGroupCNC = groupOrdersN1_CNC.Where(c => c.gKey.Equals(lastOrderBWCNCHob10.TeethShape)).ToList();
                        List<OrderModel> orders = new List<OrderModel>();
                        ordersFromGroupCNC.ForEach(c => orders.AddRange(c.gItems.ToList()));
                        var sortedOrders = SortOrders(orders, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        //refresh ListOrdersBWCNCHob10
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode( sortedOrders, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //remove group đã chia cho BWCNCHob10
                        groupOrdersN1_CNC = groupOrdersN1_CNC.Except(ordersFromGroupCNC).ToList();
                    }
                }



            }

            //2020-05-28 Start add: tạm thời khóa lại, chỉnh lại cách chia đơn N+1 ra các máy BW và CNC trên màn hình của BW
            //string htmlBW11 = BWCNCViewModel_Sub.Instance.HTMLTableOrders(ListOrdersBWHob11.ToList(), "BW11");
            //    string htmlBW12 = BWCNCViewModel_Sub.Instance.HTMLTableOrders(ListOrdersBWHob12.ToList(), "BW12");
            //    string htmlBWCNC9 = BWCNCViewModel_Sub.Instance.HTMLTableOrders(ListOrdersBWCNCHob9.ToList(), "BWCNC9");
            //    string htmlBWCNC10 = BWCNCViewModel_Sub.Instance.HTMLTableOrders(ListOrdersBWCNCHob10.ToList(), "BWCNC10");
            //    mailBody.AppendLine(htmlBW11);
            //    mailBody.AppendLine(htmlBW12);
            //    mailBody.AppendLine(htmlBWCNC9);
            //    mailBody.AppendLine(htmlBWCNC10);

            //string newSourceJson = FormatHTMLOrders(lstOrdersN1);
            //mailBody.AppendLine("Có orders " + orderType + " sau sắp xếp</br>");
            //mailBody.AppendLine("-------------------------------------------------- </br>");
            //mailBody.AppendLine(newSourceJson + "</br>");
            //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW               
            //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"

            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Sắp xếp ưu tiên " + orderType, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn");
            
            //2020-02-27 End add    

            //2020-05-28 End add


            //2020-02-28 Start lock
            //try
            //{
            //    OrderBLL.Instance.Update_bwIsSetup(status);
            //}
            //catch (Exception ex)
            //{
            //    throw new Exception("Cập nhật trạng thái orders N+1 thất bại\n" + ex.Message);
            //}
            //2020-02-28


        }
          
        //function chia đơn BW -> CNC
        private void SharingBWOrdersToCNC()
        {
            var t = Convert.ToInt32(DateTime.Now.ToString("HH"));
            StringBuilder mailBody = new StringBuilder();
            //2020-02-24 Start edit:  
            //remove: //nếu thời gian hiện tại > 14:00 hoặc < 6:00 thì các đơn có N+1 sẽ ko chuyển qua máy CNC
            //add: 
            int dayToAddN1 = HolidayHelper.Instance.DayCountN1;
            if (t>=6 && t < 24)  //vẫn nằm trong ngày hiện tại => lấy các orders có ngày xuất bằng ngày hiện tại + 1
            {
                //2019-11-25 Start add: nếu BWCNCHob9 hoặc BWCNCHob10 hết đơn thì đẩy bớt từ BW qua: SL Đơn hàng > 30 và cùng shape
                if (ListOrdersBWCNCHob9 == null)
                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>();

                if (ListOrdersBWCNCHob9 != null)
                {
                    //Nếu máy BWCNCHob9 hết đơn
                    if (ListOrdersBWCNCHob9.Count == 0)
                    {
                        //2020-02-08 Start edit: gán giá trị cho BWCNCHob9_NeedOrders trên server
                        //BWCNCHob9_NeedOrders = true;
                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true

                        //lấy BWCNCHob9_NeedOrders từ server 
                        var bw_cnc9_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob9();
                        //2020-02-08 End edit

                        //lấy từ BWHob11 hoặc BWHob12 qua
                        if (bw_cnc9_HasBWOrders == 1)
                        {
                            if (ListOrdersBWHob11 == null)
                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                            if (ListOrdersBWHob12 == null)
                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                            //2019-12-16 Start add: tách lấy các đơn N+1 ra
                            var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();
                            var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();
                            //2019-12-16 End add

                            //chỉ tách lấy các orders ko nằm trong group các orders N+1                            
                            var orderBWHob11_group = ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                            var orderBWHob12_group = ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });

                            //nếu số lương group > 1
                            if (orderBWHob11_group.Count() > 1)
                            {
                                List<OrderModel> lst = new List<OrderModel>();
                                //nếu ordersBWHob11_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                                var orderBWHob11_GroupItems = ordersBWHob11_N1.Count == 0 ? orderBWHob11_group.Skip(1) : orderBWHob11_group;
                                foreach (var g in orderBWHob11_GroupItems)//bỏ group đầu nếu ordersBWHob11_N1.Count == 0
                                {
                                    if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                    {
                                        //kiểm tra nếu BWCNCHob9 có TeethShape giống nhau
                                        if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                        {
                                            //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                            var ordersForBWCNCHob9 = ListOrdersBWHob11.Except(ordersBWHob11_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                            lst.AddRange(ordersForBWCNCHob9);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (ordersForBWCNCHob9 != null && ordersForBWCNCHob9.Count() > 0)
                                            {                                                
                                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(ordersForBWCNCHob9));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                        else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob9 (bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                        {
                                            var items = ordersBWHob11_N1.Count == 0 ? ListOrdersBWHob11.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                            lst.AddRange(items);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (items != null && items.Count() > 0)
                                            {
                                                
                                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(items));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }


                                    }
                                }

                                //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                                var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(orderedBWCNCHob9);
                                //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                //BWCNCHob9_NeedOrders = false;
                                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false
                                //2020-02-08 End edit

                            }

                            //nếu ListOrdersBWCNCHob9 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                            if (ListOrdersBWCNCHob9.Count == 0 && orderBWHob12_group.Count() > 1)
                            {
                                List<OrderModel> lst = new List<OrderModel>();
                                //nếu ordersBWHob12_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                                var orderBWHob12_GroupItems = ordersBWHob12_N1.Count == 0 ? orderBWHob12_group.Skip(1) : orderBWHob12_group;
                                foreach (var g in orderBWHob12_GroupItems)//bỏ group đầu nếu ordersBWHob12_N1.Count == 0
                                {
                                    if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                    {
                                        //kiểm tra nếu BWCNCHob9 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                        if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                        {
                                            //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                            var _ordersForBWCNCHob9 = ListOrdersBWHob12.Except(ordersBWHob12_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                            lst.AddRange(_ordersForBWCNCHob9);
                                            //remove các items sau khi đây từ BWHob12 qua
                                            if (_ordersForBWCNCHob9 != null && _ordersForBWCNCHob9.Count() > 0)
                                            {
                                                
                                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(_ordersForBWCNCHob9));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                        else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                        {
                                            var items = ordersBWHob12_N1.Count == 0 ? ListOrdersBWHob12.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                            lst.AddRange(items);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (items != null && items.Count() > 0)
                                            {
                                                
                                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(items));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                    }
                                }
                                //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                                var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(orderedBWCNCHob9);
                                //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                //BWCNCHob9_NeedOrders = false;
                                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false
                                //2020-02-08 End edit

                            }
                        }


                    }

                    Calc_PO_and_Orders_BW();
                }

                //Nếu máy BWCNCHob10 hết đơn
                if (ListOrdersBWCNCHob10 == null)
                    ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>();

                if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count == 0)
                {
                    //2020-02-08 Start edit: gán giá trị cho BWCNCHob10_NeedOrders trên server
                    //BWCNCHob10_NeedOrders = true;
                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true

                    //lấy BWCNCHob10_NeedOrders từ server 
                    var bw_cnc10_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob10();
                    //2020-02-08 End edit

                    //lấy từ BWHob11 hoặc BWHob12 qua
                    if (bw_cnc10_HasBWOrders == 1)
                    {
                        if (ListOrdersBWHob11 == null)
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                        if (ListOrdersBWHob12 == null)
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                        //2019-12-16 Start add: tách lấy các đơn N+1 ra
                        var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();
                        var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();
                        //2019-12-16 End add

                        //chỉ tách lấy các orders ko nằm trong group các orders N+1                            
                        var orderBWHob11_group = ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                        var orderBWHob12_group = ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });

                      
                        //nếu số lương group > 1
                        if (orderBWHob11_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob11_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                            var orderBWHob11_GroupItems = ordersBWHob11_N1.Count == 0 ? orderBWHob11_group.Skip(1) : orderBWHob11_group;
                            foreach (var g in orderBWHob11_GroupItems)
                            {
                                if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                {
                                    //kiểm tra nếu BWCNCHob10 có TeethShape giống TeethShape của item trong group
                                    if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob11 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob9
                                        var ordersForBWCNCHob10 = ListOrdersBWHob11.Except(ordersBWHob11_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(ordersForBWCNCHob10);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (ordersForBWCNCHob10 != null && ordersForBWCNCHob10.Count() > 0)
                                        {
                                           
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(ordersForBWCNCHob10));
                                        }
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob10 ,bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                    {
                                        var items = ordersBWHob11_N1.Count == 0 ? ListOrdersBWHob11.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {                                            
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(items));
                                        }
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }


                                }
                            }

                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                            var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(orderedBWCNCHob10);
                            //2020-02-08 Start edit:
                            //BWCNCHob10_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false
                            //2020-02-08 End edit

                        }

                        //nếu ListOrdersBWCNCHob10 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                        if (ListOrdersBWCNCHob10.Count == 0 && orderBWHob12_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob12_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                            var orderBWHob12_GroupItems = ordersBWHob12_N1.Count == 0 ? orderBWHob12_group.Skip(1) : orderBWHob12_group;
                            foreach (var g in orderBWHob12_group.Skip(1))
                            {
                                if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                {
                                    //kiểm tra nếu BWCNCHob10 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                    if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob12 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob10
                                        var _ordersForBWCNCHob10 = ListOrdersBWHob12.Except(ordersBWHob12_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(_ordersForBWCNCHob10);
                                        //remove các items sau khi đây từ BWHob12 qua
                                        if (_ordersForBWCNCHob10 != null && _ordersForBWCNCHob10.Count() > 0)
                                        {
                                            
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(_ordersForBWCNCHob10));
                                        }
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                    {
                                        var items = ordersBWHob12_N1.Count == 0 ? ListOrdersBWHob12.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {
                                            
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(items));
                                        }
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                }
                            }
                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob10
                            var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(orderedBWCNCHob10);
                            //2020-02-08 Start edit:
                            //BWCNCHob10_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false
                            //2020-02-08 End edit

                        }
                        Calc_PO_and_Orders_BW();
                    }


                }
                //2019-11-25 End add
            }
            else if(t >=0 && t < 6) //qua ngày hôm sau => lấy các orders có ngày xuất bằng ngày hiện tại
            {
                //2019-11-25 Start add: nếu BWCNCHob9 hoặc BWCNCHob10 hết đơn thì đẩy bớt từ BW qua: SL Đơn hàng > 30 và cùng shape
                if (ListOrdersBWCNCHob9 == null)
                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>();

                if (ListOrdersBWCNCHob9 != null)
                {
                    //Nếu máy BWCNCHob9 hết đơn
                    if (ListOrdersBWCNCHob9.Count == 0)
                    {
                        //2020-02-08 Start edit: gán giá trị cho BWCNCHob9_NeedOrders trên server
                        //BWCNCHob9_NeedOrders = true;
                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true

                        //lấy BWCNCHob9_NeedOrders từ server 
                        var bw_cnc9_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob9();
                        //2020-02-08 End edit
                        //lấy từ BWHob11 hoặc BWHob12 qua
                        if (bw_cnc9_HasBWOrders == 1)
                        {
                            if (ListOrdersBWHob11 == null)
                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                            if (ListOrdersBWHob12 == null)
                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                            //2019-12-16 Start add: tách lấy các đơn N+1 ra
                            var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                            var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                            //2019-12-16 End add

                            //chỉ tách lấy các orders ko nằm trong group các orders N+1                            
                            var orderBWHob11_group = ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                            var orderBWHob12_group = ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });

                            //nếu số lương group > 1
                            if (orderBWHob11_group.Count() > 1)
                            {
                                List<OrderModel> lst = new List<OrderModel>();
                                //nếu ordersBWHob11_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                                var orderBWHob11_GroupItems = ordersBWHob11_N1.Count == 0 ? orderBWHob11_group.Skip(1) : orderBWHob11_group;
                                foreach (var g in orderBWHob11_GroupItems)//bỏ group đầu nếu ordersBWHob11_N1.Count == 0
                                {
                                    if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                    {
                                        //kiểm tra nếu BWCNCHob9 có TeethShape giống nhau
                                        if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                        {
                                            //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                            var ordersForBWCNCHob9 = ListOrdersBWHob11.Except(ordersBWHob11_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                            lst.AddRange(ordersForBWCNCHob9);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (ordersForBWCNCHob9 != null && ordersForBWCNCHob9.Count() > 0)
                                            {

                                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(ordersForBWCNCHob9));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                        else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob9 (bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                        {
                                            var items = ordersBWHob11_N1.Count == 0 ? ListOrdersBWHob11.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                            lst.AddRange(items);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (items != null && items.Count() > 0)
                                            {

                                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(items));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }


                                    }
                                }

                                //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                                var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(orderedBWCNCHob9);
                                //2020-02-08 Start edit: gán giá trị BWCNCHob9_NeedOrders trên server
                                //BWCNCHob9_NeedOrders = false;
                                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false   
                                //2020-02-08 End edit

                            }

                            //nếu ListOrdersBWCNCHob9 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                            if (ListOrdersBWCNCHob9.Count == 0 && orderBWHob12_group.Count() > 1)
                            {
                                List<OrderModel> lst = new List<OrderModel>();
                                //nếu ordersBWHob12_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                                var orderBWHob12_GroupItems = ordersBWHob12_N1.Count == 0 ? orderBWHob12_group.Skip(1) : orderBWHob12_group;
                                foreach (var g in orderBWHob12_GroupItems)//bỏ group đầu nếu ordersBWHob12_N1.Count == 0
                                {
                                    if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                    {
                                        //kiểm tra nếu BWCNCHob9 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                        if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                        {
                                            //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                            var _ordersForBWCNCHob9 = ListOrdersBWHob12.Except(ordersBWHob12_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                            lst.AddRange(_ordersForBWCNCHob9);
                                            //remove các items sau khi đây từ BWHob12 qua
                                            if (_ordersForBWCNCHob9 != null && _ordersForBWCNCHob9.Count() > 0)
                                            {

                                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(_ordersForBWCNCHob9));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                        else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                        {
                                            var items = ordersBWHob12_N1.Count == 0 ? ListOrdersBWHob12.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                            lst.AddRange(items);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (items != null && items.Count() > 0)
                                            {

                                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(items));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                    }
                                }
                                //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                                var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(orderedBWCNCHob9);
                                //2020-02-08 Start edit: gán giá trị BWCNCHob9_NeedOrders trên server
                                //BWCNCHob9_NeedOrders = false;
                                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false   
                                //2020-02-08 End edit

                            }
                        }


                    }

                    Calc_PO_and_Orders_BW();
                }

                //Nếu máy BWCNCHob10 hết đơn
                if (ListOrdersBWCNCHob10 == null)
                    ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>();

                if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count == 0)
                {
                    //2020-02-08 Start edit: gán giá trị cho BWCNCHob10_NeedOrders trên server
                    //BWCNCHob10_NeedOrders = true;
                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true

                    //lấy BWCNCHob10_NeedOrders từ server 
                    var bw_cnc10_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob10();
                    //2020-02-08 End edit
                    //lấy từ BWHob11 hoặc BWHob12 qua
                    if (bw_cnc10_HasBWOrders == 1)
                    {
                        if (ListOrdersBWHob11 == null)
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                        if (ListOrdersBWHob12 == null)
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                        //2019-12-16 Start add: tách lấy các đơn N+1 ra
                        var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                        var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                        //2019-12-16 End add

                        //chỉ tách lấy các orders ko nằm trong group các orders N+1                            
                        var orderBWHob11_group = ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                        var orderBWHob12_group = ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });


                        //nếu số lương group > 1
                        if (orderBWHob11_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob11_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                            var orderBWHob11_GroupItems = ordersBWHob11_N1.Count == 0 ? orderBWHob11_group.Skip(1) : orderBWHob11_group;
                            foreach (var g in orderBWHob11_GroupItems)
                            {
                                if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                {
                                    //kiểm tra nếu BWCNCHob10 có TeethShape giống TeethShape của item trong group
                                    if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob11 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob9
                                        var ordersForBWCNCHob10 = ListOrdersBWHob11.Except(ordersBWHob11_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(ordersForBWCNCHob10);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (ordersForBWCNCHob10 != null && ordersForBWCNCHob10.Count() > 0)
                                        {

                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(ordersForBWCNCHob10));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob10 ,bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                    {
                                        var items = ordersBWHob11_N1.Count == 0 ? ListOrdersBWHob11.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(items));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }


                                }
                            }

                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                            var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(orderedBWCNCHob10);

                            //2020-02-08 Start edit: gán giá trị BWCNCHob10_NeedOrders trên server
                            //BWCNCHob10_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false   
                            //2020-02-08 End edit

                        }

                        //nếu ListOrdersBWCNCHob10 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                        if (ListOrdersBWCNCHob10.Count == 0 && orderBWHob12_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob12_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                            var orderBWHob12_GroupItems = ordersBWHob12_N1.Count == 0 ? orderBWHob12_group.Skip(1) : orderBWHob12_group;
                            foreach (var g in orderBWHob12_group.Skip(1))
                            {
                                if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                {
                                    //kiểm tra nếu BWCNCHob10 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                    if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob12 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob10
                                        var _ordersForBWCNCHob10 = ListOrdersBWHob12.Except(ordersBWHob12_N1).Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(_ordersForBWCNCHob10);
                                        //remove các items sau khi đây từ BWHob12 qua
                                        if (_ordersForBWCNCHob10 != null && _ordersForBWCNCHob10.Count() > 0)
                                        {

                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(_ordersForBWCNCHob10));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                    {
                                        var items = ordersBWHob12_N1.Count == 0 ? ListOrdersBWHob12.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList() : ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {

                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(items));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                }
                            }
                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob10
                            var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(orderedBWCNCHob10);
                            //2020-02-08 Start edit: gán giá trị BWCNCHob10_NeedOrders trên server
                            //BWCNCHob10_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false   
                            //2020-02-08 End edit

                        }
                        Calc_PO_and_Orders_BW();
                    }


                }
                //2019-11-25 End add
            }
            else
            {
                //2019-11-25 Start add: nếu BWCNCHob9 hoặc BWCNCHob10 hết đơn thì đẩy bớt từ BW qua: SL Đơn hàng > 30 và cùng shape
                if (ListOrdersBWCNCHob9 == null)
                    ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>();

                if (ListOrdersBWCNCHob9 != null)
                {
                    //Nếu máy BWCNCHob9 hết đơn
                    if (ListOrdersBWCNCHob9.Count == 0)
                    {
                        //2020-02-08 Start edit: gán giá trị cho BWCNCHob9_NeedOrders trên server
                        //BWCNCHob9_NeedOrders = true;
                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true

                        //lấy BWCNCHob9_NeedOrders từ server 
                        var bw_cnc9_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob9();
                        //2020-02-08 End edit
                        //lấy từ BWHob11 hoặc BWHob12 qua
                        if (bw_cnc9_HasBWOrders == 1)
                        {
                            if (ListOrdersBWHob11 == null)
                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                            if (ListOrdersBWHob12 == null)
                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                          

                            var orderBWHob11_group = ListOrdersBWHob11.GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                            var orderBWHob12_group = ListOrdersBWHob12.GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });

                            //nếu số lương group > 1
                            if (orderBWHob11_group.Count() > 1)
                            {
                                List<OrderModel> lst = new List<OrderModel>();
                                foreach (var g in orderBWHob11_group.Skip(1))//bỏ group đầu
                                {
                                    if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                    {
                                        //kiểm tra nếu BWCNCHob9 có TeethShape giống nhau
                                        if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                        {
                                            //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                            var ordersForBWCNCHob9 = ListOrdersBWHob11.Where(c => c.TeethShape.Equals(g.teethShape));
                                            lst.AddRange(ordersForBWCNCHob9);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (ordersForBWCNCHob9 != null && ordersForBWCNCHob9.Count() > 0)
                                            {                                                
                                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(lst));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                        else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob9 (bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                        {
                                            var items = ListOrdersBWHob11.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                            lst.AddRange(items);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (items != null && items.Count() > 0)
                                            {                                                
                                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(lst));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }


                                    }
                                }

                                //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                                var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(orderedBWCNCHob9);
                                //2020-02-08 Start edit: gán giá trị BWCNCHob9_NeedOrders trên server
                                //BWCNCHob9_NeedOrders = false;
                                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false   
                                //2020-02-08 End edit

                            }

                            //nếu ListOrdersBWCNCHob9 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                            if (ListOrdersBWCNCHob9.Count == 0 && orderBWHob12_group.Count() > 1)
                            {
                                List<OrderModel> lst = new List<OrderModel>();
                                foreach (var g in orderBWHob12_group.Skip(1))//bỏ group đầu
                                {
                                    if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                    {
                                        //kiểm tra nếu BWCNCHob9 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                        if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                        {
                                            //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                            var _ordersForBWCNCHob9 = ListOrdersBWHob12.Where(c => c.TeethShape.Equals(g.teethShape));
                                            lst.AddRange(_ordersForBWCNCHob9);
                                            //remove các items sau khi đây từ BWHob12 qua
                                            if (_ordersForBWCNCHob9 != null && _ordersForBWCNCHob9.Count() > 0)
                                            {                                               
                                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(lst));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                        else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                        {
                                            var items = ListOrdersBWHob12.GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                            lst.AddRange(items);
                                            //remove các items sau khi đây từ BWHob11 qua
                                            if (items != null && items.Count() > 0)
                                            {                                                
                                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(lst));
                                            }
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                            //2020-02-08 End edit
                                            break;//chỉ đẩy 1 group
                                        }
                                    }
                                }
                                //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                                var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(orderedBWCNCHob9);
                                //2020-02-08 Start edit: gán giá trị BWCNCHob9_NeedOrders trên server
                                //BWCNCHob9_NeedOrders = false;
                                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false   
                                //2020-02-08 End edit

                            }
                        }


                    }

                    Calc_PO_and_Orders_BW();
                }

                //Nếu máy BWCNCHob10 hết đơn
                if (ListOrdersBWCNCHob10 == null)
                    ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>();

                if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count == 0)
                {
                    //2020-02-08 Start edit: gán giá trị cho BWCNCHob10_NeedOrders trên server
                    //BWCNCHob10_NeedOrders = true;
                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true

                    //lấy BWCNCHob10_NeedOrders từ server 
                    var bw_cnc10_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob10();
                    //2020-02-08 End edit
                    //lấy từ BWHob11 hoặc BWHob12 qua
                    if (bw_cnc10_HasBWOrders == 1)
                    {
                        if (ListOrdersBWHob11 == null)
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                        if (ListOrdersBWHob12 == null)
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                        var orderBWHob11_group = ListOrdersBWHob11.Except(ListOrdersBWCNCHob9).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                        var orderBWHob12_group = ListOrdersBWHob12.Except(ListOrdersBWCNCHob9).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });

                        //nếu số lương group > 1
                        if (orderBWHob11_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            foreach (var g in orderBWHob11_group.Skip(1))//bỏ group đầu
                            {
                                if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                {
                                    //kiểm tra nếu BWCNCHob10 có TeethShape giống TeethShape của item trong group
                                    if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob11 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob9
                                        var ordersForBWCNCHob10 = ListOrdersBWHob11.Except(ListOrdersBWCNCHob9).Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(ordersForBWCNCHob10);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (ordersForBWCNCHob10 != null && ordersForBWCNCHob10.Count() > 0)
                                        {                                            
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(lst));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob10 ,bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                    {
                                        var items = ListOrdersBWHob11.Except(ListOrdersBWCNCHob9).GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {                                            
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ListOrdersBWHob11.Except(lst));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }


                                }
                            }

                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                            var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(orderedBWCNCHob10);
                            //2020-02-08 Start edit: gán giá trị BWCNCHob10_NeedOrders trên server
                            //BWCNCHob10_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false   
                            //2020-02-08 End edit

                        }

                        //nếu ListOrdersBWCNCHob10 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                        if (ListOrdersBWCNCHob10.Count == 0 && orderBWHob12_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            foreach (var g in orderBWHob12_group.Skip(1))//bỏ group đầu
                            {
                                if (g.number_of_orders > 30)//nếu SL đơn hàng > 30
                                {
                                    //kiểm tra nếu BWCNCHob10 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                    if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob12 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob10
                                        var _ordersForBWCNCHob10 = ListOrdersBWHob12.Except(ListOrdersBWCNCHob9).Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(_ordersForBWCNCHob10);
                                        //remove các items sau khi đây từ BWHob12 qua
                                        if (_ordersForBWCNCHob10 != null && _ordersForBWCNCHob10.Count() > 0)
                                        {                                            
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(lst));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                    {
                                        var items = ListOrdersBWHob12.Except(ListOrdersBWCNCHob9).GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {                                            
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ListOrdersBWHob12.Except(lst));
                                        }
                                        //bật flags BWCNCHob10_HasOrders                                     
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasOrders trên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                        //2020-02-08 End edit
                                        break;//chỉ đẩy 1 group
                                    }
                                }
                            }
                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob10
                            var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                            ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(orderedBWCNCHob10);
                            //2020-02-08 Start edit: gán giá trị BWCNCHob10_NeedOrders trên server
                            //BWCNCHob10_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false   
                            //2020-02-08 End edit

                        }
                        Calc_PO_and_Orders_BW();
                    }


                }
                //2019-11-25 End add
            }
        }

        private void SharingBWOrdersToCNC_NormalMode()
        {
            var t = Convert.ToInt32(DateTime.Now.ToString("HH"));
            StringBuilder mailBody = new StringBuilder();
            //2020-02-24 Start edit:  
            //remove: //nếu thời gian hiện tại > 14:00 hoặc < 6:00 thì các đơn có N+1 sẽ ko chuyển qua máy CNC
            
            //từ 6:00 đến 23:59 => chỉ lấy ra N+1
            if (t >= 6 && t < 24)  //vẫn nằm trong ngày hiện tại => lấy các orders có ngày xuất bằng ngày hiện tại + 1
            {
                ShareBWOrders_Normal6To24();
            }else if(t>=0 && t<6) //từ 0:00 đến 5:59 => lấy ra đơn N
            {
                ShareBWOrders_Normal0To6();
            }
           
        }

        private void SharingBWOrdersToCNC_SSDMode()
        {
            var t = Convert.ToInt32(DateTime.Now.ToString("HH"));
            StringBuilder mailBody = new StringBuilder();
            //2020-02-24 Start edit:  
            //remove: //nếu thời gian hiện tại > 14:00 hoặc < 6:00 thì các đơn có N+1 sẽ ko chuyển qua máy CNC

            //từ 6:00 đến 23:59 => chỉ lấy ra N+1
            if (t >= 6 && t < 24)  //vẫn nằm trong ngày hiện tại => lấy các orders có ngày xuất bằng ngày hiện tại + 1
            {
                ShareBWOrders_Normal6To24();//2020-03-21 test nếu trong SSDMode mà máy CNC hết order vẫn chuyển đơn qua
            }
            else if (t >= 0 && t < 6) //từ 0:00 đến 5:59 => lấy ra đơn N
            {
                ShareBWOrders_Normal0To6();//2020-03-21 test nếu trong SSDMode mà máy CNC hết order vẫn chuyển đơn qua
            }

        }
        //2020-03-12 Start add: tách chức năng chia đơn BW cho CNC
        private void ShareBWOrders_Normal6To24()
        {
            StringBuilder mailBody;
            int dayAddToCheckN1 = HolidayHelper.Instance.DayCountN1; //DateTime.Now.ToString("dddd").Equals("Saturday") ? 2 : 1;//nếu là thứ 7 thì N+1 +2
            //2019-11-25 Start add: nếu BWCNCHob9 hoặc BWCNCHob10 hết đơn thì đẩy bớt từ BW qua: SL Đơn hàng > 30 và cùng shape
            if (ListOrdersBWCNCHob9 == null)
                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>();

            if (ListOrdersBWCNCHob9 != null)
            {
                //Nếu máy BWCNCHob9 hết đơn
                if (ListOrdersBWCNCHob9.Count == 0)
                {
                    mailBody = new StringBuilder();

                    //2020-02-08 Start edit: gán giá trị cho BWCNCHob9_NeedOrders trên server
                    //BWCNCHob9_NeedOrders = true;
                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true

                    //lấy BWCNCHob9_NeedOrders từ server 
                    var bw_cnc9_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob9();
                    //2020-02-08 End edit

                    //lấy từ BWHob11 hoặc BWHob12 qua
                    if (bw_cnc9_HasBWOrders == 1)
                    {
                        if (ListOrdersBWHob11 == null)
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                        if (ListOrdersBWHob12 == null)
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                        //2019-12-16 Start add: tách lấy các đơn N+1 ra
                        //var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();
                        //var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();
                        //2019-12-16 End add


                        var orderBWHob11_group = ListOrdersBWHob11.GroupBy(c =>  c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                        var orderBWHob12_group = ListOrdersBWHob12.GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });

                        //nếu số lương group > 1
                        if (orderBWHob11_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob11_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                            var orderBWHob11_GroupItems = orderBWHob11_group.Where(c => !c.teethShape.Equals(Hob11BW_TeethShape)).Skip(1);
                            foreach (var g in orderBWHob11_GroupItems)//bỏ group đầu nếu ordersBWHob11_N1.Count == 0
                            {
                                if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                                {
                                    //kiểm tra nếu BWCNCHob9 có TeethShape giống nhau
                                    if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                        var ordersForBWCNCHob9 = ListOrdersBWHob11.Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(ordersForBWCNCHob9);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (ordersForBWCNCHob9 != null && ordersForBWCNCHob9.Count() > 0)
                                        {
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(ordersForBWCNCHob9).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob9 (bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                    {
                                        var items =ListOrdersBWHob11.Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c =>c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {

                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }


                                }
                                else if (g.teethShape == orderBWHob11_GroupItems.LastOrDefault().teethShape)
                                {
                                    //nếu đây là group cuối
                                    //đẩy group cuối qua cho BWCNC9
                                    var items = ListOrdersBWHob11.Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                        //bật flags BWCNCHob9_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                        //BWCNCHob9_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                //2020-02-08 End edit
                                    }

                                }

                                
                            }

                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                            //var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            var orderedBWCNCHob9 = SetupSSD_BW(lst, SSDType.N1,Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));                            
                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                            //BWCNCHob9_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false
                                                                                    //2020-02-08 End edit
                            if (orderedBWCNCHob9 != null && orderedBWCNCHob9.Count > 0)
                            {
                                string sourceJson = FormatHTMLOrders(ListOrdersBWHob11.ToList());
                                string newSourceJson = FormatHTMLOrders(lst);
                                mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob9.Count + " orders qua máy CNC9 (A1)</br>");
                                mailBody.AppendLine("----------------------------------------------------------------------</br>");
                                mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong></br>");
                                mailBody.AppendLine(sourceJson + "</br>");
                                mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong></br>");
                                mailBody.AppendLine(newSourceJson + "</br>");
                                //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                                //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW11)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                                //2020-02-27 End add
                            }

                        }

                        //nếu ListOrdersBWCNCHob9 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                        if (ListOrdersBWCNCHob9.Count == 0 && orderBWHob12_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob12_N1.Count == 0 và ordersBWHob12_N.count == 0 => ko có item nào N+1, N => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                            var orderBWHob12_GroupItems = orderBWHob12_group.Where(c=> !c.teethShape.Equals(Hob12BW_TeethShape)).Skip(1);
                            foreach (var g in orderBWHob12_GroupItems)//bỏ group đầu nếu ordersBWHob12_N1.Count == 0
                            {
                                if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                                {
                                    //kiểm tra nếu BWCNCHob9 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                    if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                        var _ordersForBWCNCHob9 = ListOrdersBWHob12.Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(_ordersForBWCNCHob9);
                                        //remove các items sau khi đây từ BWHob12 qua
                                        if (_ordersForBWCNCHob9 != null && _ordersForBWCNCHob9.Count() > 0)
                                        {
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(_ordersForBWCNCHob9).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                    {
                                        var items = ListOrdersBWHob12.Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c => c.TeethShape).Skip(1).FirstOrDefault().ToList(); 
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }
                                }else if (g.teethShape == orderBWHob12_GroupItems.LastOrDefault().teethShape)
                                {
                                    //nếu đây là group cuối
                                    //đẩy group cuối qua cho BWCNC9
                                    var items =  ListOrdersBWHob12.Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                        //bật flags BWCNCHob9_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                        //BWCNCHob9_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                //2020-02-08 End edit
                                    }

                                }

                            }


                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                            //var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            var orderedBWCNCHob9 = SetupSSD_BW(lst, SSDType.N1,Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));

                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                            //BWCNCHob9_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false
                                                                                    //2020-02-08 End edit
                            if (orderedBWCNCHob9 != null && orderedBWCNCHob9.Count > 0)
                            {
                                string sourceJson = FormatHTMLOrders(ListOrdersBWHob12.ToList());
                                string newSourceJson = FormatHTMLOrders(lst);
                                mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob9.Count + " orders qua máy CNC9 (A2)</br>");
                                mailBody.AppendLine("----------------------------------------------------------------------</br>");
                                mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong></br>");
                                mailBody.AppendLine(sourceJson + "</br>");
                                mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong></br>");
                                mailBody.AppendLine(newSourceJson + "</br>");
                                //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                                //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW12)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                                //2020-02-27 End add
                            }

                        }
                    }


                }

                Calc_PO_and_Orders_BW();
            }



            //Nếu máy BWCNCHob10 hết đơn
            if (ListOrdersBWCNCHob10 == null)
                ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>();

            if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count == 0)
            {
                mailBody = new StringBuilder();

                //2020-02-08 Start edit: gán giá trị cho BWCNCHob10_NeedOrders trên server
                //BWCNCHob10_NeedOrders = true;
                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true

                //lấy BWCNCHob10_NeedOrders từ server 
                var bw_cnc10_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob10();
                //2020-02-08 End edit

                //lấy từ BWHob11 hoặc BWHob12 qua
                if (bw_cnc10_HasBWOrders == 1)
                {
                    if (ListOrdersBWHob11 == null)
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                    if (ListOrdersBWHob12 == null)
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                    //2019-12-16 Start add: tách lấy các đơn N+1 ra
                    //var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();
                    //var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();
                    //2019-12-16 End add


                    //chỉ tách lấy các orders ko nằm trong group các orders N+1 và N                            
                    var orderBWHob11_group = ListOrdersBWHob11.GroupBy(c =>c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                    var orderBWHob12_group = ListOrdersBWHob12.GroupBy(c =>c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });


                    //nếu số lương group > 1
                    if (orderBWHob11_group.Count() > 1)
                    {
                        List<OrderModel> lst = new List<OrderModel>();
                        //nếu ordersBWHob11_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                        var orderBWHob11_GroupItems = orderBWHob11_group.Where(c => !c.teethShape.Equals(Hob11BW_TeethShape)).Skip(1);
                        foreach (var g in orderBWHob11_GroupItems)
                        {
                            if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                            {
                                //kiểm tra nếu BWCNCHob10 có TeethShape giống TeethShape của item trong group
                                if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                {
                                    //đẩy group orders của BWHob11 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob9
                                    var ordersForBWCNCHob10 = ListOrdersBWHob11.Except(ListOrdersBWCNCHob9.ToList()).Where(c => c.TeethShape.Equals(g.teethShape));
                                    lst.AddRange(ordersForBWCNCHob10);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (ordersForBWCNCHob10 != null && ordersForBWCNCHob10.Count() > 0)
                                    {
                                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(ordersForBWCNCHob10).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }
                                else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob10 ,bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                {
                                    var items =ListOrdersBWHob11.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c =>c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }


                            }
                            else if (g.teethShape == orderBWHob11_GroupItems.LastOrDefault().teethShape)
                            {
                                //nếu đây là group cuối
                                //đẩy group cuối qua cho BWCNC10
                                var items =  ListOrdersBWHob11.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                lst.AddRange(items);
                                //remove các items sau khi đây từ BWHob11 qua
                                if (items != null && items.Count() > 0)
                                {
                                    ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(),Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                    //bật flags BWCNCHob10_HasOrders
                                    //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                    //BWCNCHob9_HasBWOrders = true;
                                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                             //2020-02-08 End edit
                                }
                                
                            }

                        }

                        //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                        //var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        var orderedBWCNCHob10 = SetupSSD_BW(lst, SSDType.N1,Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob10, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //2020-02-08 Start edit:
                        //BWCNCHob10_NeedOrders = false;
                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false
                                                                                 //2020-02-08 End edit
                        if (orderedBWCNCHob10 != null && orderedBWCNCHob10.Count > 0)
                        {
                            string sourceJson = FormatHTMLOrders(ListOrdersBWHob11.ToList());
                            string newSourceJson = FormatHTMLOrders(lst);
                            mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob10.Count + " orders qua máy CNC10 (B1) </br>");
                            mailBody.AppendLine("----------------------------------------------------------------------</br>");
                            mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong> </br>");
                            mailBody.AppendLine(sourceJson + "</br>");
                            mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong> </br>");
                            mailBody.AppendLine(newSourceJson + "</br>");
                            //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                            //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW11)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                            //2020-02-27 End add
                        }
                    }

                    //nếu ListOrdersBWCNCHob10 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                    if (ListOrdersBWCNCHob10.Count == 0 && orderBWHob12_group.Count() > 1)
                    {
                        List<OrderModel> lst = new List<OrderModel>();
                        //nếu ordersBWHob12_N1.Count == 0 => ko có item nào N+1 => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N+1 qua CNC
                        var orderBWHob12_GroupItems = orderBWHob12_group.Where(c => !c.teethShape.Equals(Hob12BW_TeethShape)).Skip(1);
                        foreach (var g in orderBWHob12_GroupItems)
                        {
                            if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                            {
                                //kiểm tra nếu BWCNCHob10 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                {
                                    //đẩy group orders của BWHob12 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob10
                                    var _ordersForBWCNCHob10 = ListOrdersBWHob12.Except(ListOrdersBWCNCHob9.ToList()).Where(c => c.TeethShape.Equals(g.teethShape));
                                    lst.AddRange(_ordersForBWCNCHob10);
                                    //remove các items sau khi đây từ BWHob12 qua
                                    if (_ordersForBWCNCHob10 != null && _ordersForBWCNCHob10.Count() > 0)
                                    {
                                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(_ordersForBWCNCHob10).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }
                                else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                {
                                    var items =  ListOrdersBWHob12.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c =>c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }
                            }
                            else if (g.teethShape == orderBWHob12_GroupItems.LastOrDefault().teethShape)
                            {
                                //nếu đây là group cuối
                                //đẩy group cuối qua cho BWCNC10
                                var items =  ListOrdersBWHob12.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                lst.AddRange(items);
                                //remove các items sau khi đây từ BWHob12 qua
                                if (items != null && items.Count() > 0)
                                {
                                    ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                    //bật flags BWCNCHob9_HasOrders
                                    //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders trên server
                                    //BWCNCHob10_HasBWOrders = true;
                                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                             //2020-02-08 End edit
                                }

                            }
                        }
                        //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob10
                        //var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        var orderedBWCNCHob10 = SetupSSD_BW(lst,SSDType.N1, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob10, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //2020-02-08 Start edit:
                        //BWCNCHob10_NeedOrders = false;
                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false
                                                                                 //2020-02-08 End edit
                        if (orderedBWCNCHob10 != null && orderedBWCNCHob10.Count > 0)
                        {
                            string sourceJson = FormatHTMLOrders(ListOrdersBWHob12.ToList());
                            string newSourceJson = FormatHTMLOrders(lst);
                            mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob10.Count + " orders qua máy CNC10 (B2)</br>");
                            mailBody.AppendLine("----------------------------------------------------------------------</br>");
                            mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong></br>");
                            mailBody.AppendLine(sourceJson + "</br>");
                            mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong></br>");
                            mailBody.AppendLine(newSourceJson + "</br>");
                            //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                            //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW12)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                            //2020-02-27 End add
                        }
                    }
                    Calc_PO_and_Orders_BW();
                }


            }
            //2019-11-25 End add

        }
        private void ShareBWOrders_Normal0To6()
        {
            StringBuilder mailBody;
            if (ListOrdersBWCNCHob9 == null)
                ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>();

            if (ListOrdersBWCNCHob9 != null)
            {
                //Nếu máy BWCNCHob9 hết đơn
                if (ListOrdersBWCNCHob9.Count == 0)
                {
                    mailBody = new StringBuilder();

                    //2020-02-08 Start edit: gán giá trị cho BWCNCHob9_NeedOrders trên server
                    //BWCNCHob9_NeedOrders = true;
                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true

                    //lấy BWCNCHob9_NeedOrders từ server 
                    var bw_cnc9_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob9();
                    //2020-02-08 End edit

                    //lấy từ BWHob11 hoặc BWHob12 qua
                    if (bw_cnc9_HasBWOrders == 1)
                    {
                        if (ListOrdersBWHob11 == null)
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                        if (ListOrdersBWHob12 == null)
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                        //2019-12-16 Start add: tách lấy các đơn N ra
                        //var ordersBWHob11_N = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                        //var ordersBWHob12_N = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                        //2019-12-16 End add


                        var orderBWHob11_group = ListOrdersBWHob11.GroupBy(c =>c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                        var orderBWHob12_group = ListOrdersBWHob12.GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });

                        //nếu số lương group > 1
                        if (orderBWHob11_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob11_N.Count == 0 => ko có item nào N => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N qua CNC
                            var orderBWHob11_GroupItems = orderBWHob11_group.Where(c => !c.teethShape.Equals(Hob11BW_TeethShape)).Skip(1);
                            foreach (var g in orderBWHob11_GroupItems)//bỏ group đầu nếu ordersBWHob11_N1.Count == 0
                            {
                                if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                                {
                                    //kiểm tra nếu BWCNCHob9 có TeethShape giống nhau
                                    if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                        var ordersForBWCNCHob9 = ListOrdersBWHob11.Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(ordersForBWCNCHob9);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (ordersForBWCNCHob9 != null && ordersForBWCNCHob9.Count() > 0)
                                        {
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(ordersForBWCNCHob9).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob9 (bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                    {
                                        var items = ListOrdersBWHob11.Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c =>c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {
                                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }
                                }
                                else if (g.teethShape == orderBWHob11_GroupItems.LastOrDefault().teethShape)
                                {
                                    //nếu đây là group cuối
                                    //đẩy group cuối qua cho BWCNC9
                                    var items = ListOrdersBWHob11.Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                        //bật flags BWCNCHob9_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                        //BWCNCHob9_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                //2020-02-08 End edit
                                    }

                                }
                            }

                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                            //var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            var orderedBWCNCHob9 = SetupSSD_BW(lst,SSDType.N1, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));
                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                            //BWCNCHob9_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false
                                                                                    //2020-02-08 End edit
                            if (orderedBWCNCHob9 != null && orderedBWCNCHob9.Count > 0)
                            {
                                string sourceJson = FormatHTMLOrders(ListOrdersBWHob11.ToList());
                                string newSourceJson = FormatHTMLOrders(lst);
                                mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob9.Count + " orders qua máy CNC9 (C1)</br>");
                                mailBody.AppendLine("----------------------------------------------------------------------</br>");
                                mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong></br>");
                                mailBody.AppendLine(sourceJson + "</br>");
                                mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong></br>");
                                mailBody.AppendLine(newSourceJson + "</br>");
                                //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                                //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW11)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                                //2020-02-27 End add
                            }

                        }

                        //nếu ListOrdersBWCNCHob9 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                        if (ListOrdersBWCNCHob9.Count == 0 && orderBWHob12_group.Count() > 1)
                        {
                            List<OrderModel> lst = new List<OrderModel>();
                            //nếu ordersBWHob12_N1.Count == 0 và ordersBWHob12_N.count == 0 => ko có item nào N, N => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N qua CNC
                            var orderBWHob12_GroupItems = orderBWHob12_group.Where(c => !c.teethShape.Equals(Hob12BW_TeethShape)).Skip(1);
                            foreach (var g in orderBWHob12_GroupItems)//bỏ group đầu nếu ordersBWHob12_N1.Count == 0
                            {
                                if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                                {
                                    //kiểm tra nếu BWCNCHob9 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                    if (g.teethShape.Equals(Hob9BWCNC_TeethShape))
                                    {
                                        //đẩy group orders của BWHob11 có teethShape giống BWCNCHob9 TeethShape qua BWCNCHob9
                                        var _ordersForBWCNCHob9 = ListOrdersBWHob12.Where(c => c.TeethShape.Equals(g.teethShape));
                                        lst.AddRange(_ordersForBWCNCHob9);
                                        //remove các items sau khi đây từ BWHob12 qua
                                        if (_ordersForBWCNCHob9 != null && _ordersForBWCNCHob9.Count() > 0)
                                        {                                        
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(_ordersForBWCNCHob9).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }
                                    else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                    {
                                        var items = ListOrdersBWHob12.Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c =>c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                        lst.AddRange(items);
                                        //remove các items sau khi đây từ BWHob11 qua
                                        if (items != null && items.Count() > 0)
                                        {
                                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                            //bật flags BWCNCHob9_HasOrders
                                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                            //BWCNCHob9_HasBWOrders = true;
                                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                    //2020-02-08 End edit
                                        }

                                        break;//chỉ đẩy 1 group
                                    }
                                }
                                else if (g.teethShape == orderBWHob12_GroupItems.LastOrDefault().teethShape)
                                {
                                    //nếu đây là group cuối
                                    //đẩy group cuối qua cho BWCNC9
                                    var items = ListOrdersBWHob12.Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob12 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                        //bật flags BWCNCHob9_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                        //BWCNCHob9_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(1);//true
                                                                                                //2020-02-08 End edit
                                    }

                                }
                            }
                            //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                            //var orderedBWCNCHob9 = SortOrders(lst, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            var orderedBWCNCHob9 = SetupSSD_BW(lst,SSDType.N1, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter);
                            ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));

                            //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                            //BWCNCHob9_NeedOrders = false;
                            MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob9(0);//false
                                                                                    //2020-02-08 End edit
                            if (orderedBWCNCHob9 != null && orderedBWCNCHob9.Count > 0)
                            {
                                string sourceJson = FormatHTMLOrders(ListOrdersBWHob12.ToList());
                                string newSourceJson = FormatHTMLOrders(lst);
                                mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob9.Count + " orders qua máy CNC9 (C2)</br>");
                                mailBody.AppendLine("----------------------------------------------------------------------</br>");
                                mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong></br>");
                                mailBody.AppendLine(sourceJson + "</br>");
                                mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong></br>");
                                mailBody.AppendLine(newSourceJson + "</br>");
                                //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                                //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW12)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                                //2020-02-27 End add
                            }

                        }
                    }


                }

                Calc_PO_and_Orders_BW();
            }

            //Nếu máy BWCNCHob10 hết đơn
            if (ListOrdersBWCNCHob10 == null)
                ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>();

            if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count == 0)
            {
                mailBody = new StringBuilder();

                //2020-02-08 Start edit: gán giá trị cho BWCNCHob10_NeedOrders trên server
                //BWCNCHob10_NeedOrders = true;
                MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true

                //lấy BWCNCHob10_NeedOrders từ server 
                var bw_cnc10_HasBWOrders = MachineTypeBLL.Instance.GetFlagCheckBWOrder_CNCHob10();
                //2020-02-08 End edit

                //lấy từ BWHob11 hoặc BWHob12 qua
                if (bw_cnc10_HasBWOrders == 1)
                {
                    if (ListOrdersBWHob11 == null)
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>();

                    if (ListOrdersBWHob12 == null)
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>();

                    //2019-12-16 Start add: tách lấy các đơn N ra
                    //var ordersBWHob11_N = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                    //var ordersBWHob12_N = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
                    //2019-12-16 End add


                    //chỉ tách lấy các orders ko nằm trong group các orders N+1 và N                            
                    var orderBWHob11_group = ListOrdersBWHob11.GroupBy(c =>c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });
                    var orderBWHob12_group = ListOrdersBWHob12.GroupBy(c =>c.TeethShape).Select(g => new { teethShape = g.Key, number_of_orders = g.Sum(c => Convert.ToInt32(c.Number_of_Orders)) });


                    //nếu số lương group > 1
                    if (orderBWHob11_group.Count() > 1)
                    {
                        List<OrderModel> lst = new List<OrderModel>();
                        //nếu ordersBWHob11_N.Count == 0 => ko có item nào N => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N qua CNC
                        var orderBWHob11_GroupItems = orderBWHob11_group.Where(c => !c.teethShape.Equals(Hob11BW_TeethShape)).Skip(1);
                        foreach (var g in orderBWHob11_GroupItems)
                        {
                            if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                            {
                                //kiểm tra nếu BWCNCHob10 có TeethShape giống TeethShape của item trong group
                                if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                {
                                    //đẩy group orders của BWHob11 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob9
                                    var ordersForBWCNCHob10 = ListOrdersBWHob11.Except(ListOrdersBWCNCHob9.ToList()).Where(c => c.TeethShape.Equals(g.teethShape));
                                    lst.AddRange(ordersForBWCNCHob10);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (ordersForBWCNCHob10 != null && ordersForBWCNCHob10.Count() > 0)
                                    {
                                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(ordersForBWCNCHob10).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }
                                else //chỉ đẩy 1 group orders của BWHob11 qua BWCNCHob10 ,bỏ qua group đầu tiên vì BWHob11 đang sử dụng
                                {
                                    var items = ListOrdersBWHob11.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c =>c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }


                            }
                            else if (g.teethShape == orderBWHob11_GroupItems.LastOrDefault().teethShape)
                            {
                                //nếu đây là group cuối
                                //đẩy group cuối qua cho BWCNC10
                                var items = ListOrdersBWHob11.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob11BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                lst.AddRange(items);
                                //remove các items sau khi đây từ BWHob11 qua
                                if (items != null && items.Count() > 0)
                                {
                                    ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob11.Except(items).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                    //bật flags BWCNCHob10_HasOrders
                                    //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders trên server
                                    //BWCNCHob9_HasBWOrders = true;
                                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                             //2020-02-08 End edit
                                }

                            }
                        }

                        //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob9
                        //var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        var orderedBWCNCHob10 = SetupSSD_BW(lst,SSDType.N1, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob10, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //2020-02-08 Start edit:
                        //BWCNCHob10_NeedOrders = false;
                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false
                                                                                 //2020-02-08 End edit
                        if (orderedBWCNCHob10 != null && orderedBWCNCHob10.Count > 0)
                        {
                            string sourceJson = FormatHTMLOrders(ListOrdersBWHob11.ToList());
                            string newSourceJson = FormatHTMLOrders(lst);
                            mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob10.Count + " orders qua máy CNC10 (D1) </br>");
                            mailBody.AppendLine("----------------------------------------------------------------------</br>");
                            mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong> </br>");
                            mailBody.AppendLine(sourceJson + "</br>");
                            mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong> </br>");
                            mailBody.AppendLine(newSourceJson + "</br>");
                            //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                            //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW11)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                            //2020-02-27 End add
                        }
                    }

                    //nếu ListOrdersBWCNCHob10 vẫn chưa có items nào và orderBWHob12_group có group items > 1
                    if (ListOrdersBWCNCHob10.Count == 0 && orderBWHob12_group.Count() > 1)
                    {
                        List<OrderModel> lst = new List<OrderModel>();
                        //nếu ordersBWHob12_N.Count == 0 => ko có item nào N => ko chuyển group items đầu của BW qua CNC, ngược lại có thể chuyển các items ko có N qua CNC
                        var orderBWHob12_GroupItems =  orderBWHob12_group.Where(c => !c.teethShape.Equals(Hob12BW_TeethShape)).Skip(1);
                        foreach (var g in orderBWHob12_GroupItems)
                        {
                            if (g.number_of_orders >= 30)//nếu SL đơn hàng >= 30
                            {
                                //kiểm tra nếu BWCNCHob10 có TeethShape giống với TeethShape của item trong ListOrdersBWHob12
                                if (g.teethShape.Equals(Hob10BWCNC_TeethShape))
                                {
                                    //đẩy group orders của BWHob12 có teethShape giống BWCNCHob10 TeethShape qua BWCNCHob10
                                    var _ordersForBWCNCHob10 = ListOrdersBWHob12.Except(ListOrdersBWCNCHob9.ToList()).Where(c => c.TeethShape.Equals(g.teethShape));
                                    lst.AddRange(_ordersForBWCNCHob10);
                                    //remove các items sau khi đây từ BWHob12 qua
                                    if (_ordersForBWCNCHob10 != null && _ordersForBWCNCHob10.Count() > 0)
                                    {
                                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(_ordersForBWCNCHob10).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }
                                else //chỉ đẩy 1 group orders của BWHob12 qua BWCNCHob9 ,bỏ qua group đầu tiên vì BWHob12 đang sử dụng
                                {
                                    var items = ListOrdersBWHob12.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c =>c.TeethShape).Skip(1).FirstOrDefault().ToList();
                                    lst.AddRange(items);
                                    //remove các items sau khi đây từ BWHob11 qua
                                    if (items != null && items.Count() > 0)
                                    {
                                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                        //bật flags BWCNCHob10_HasOrders
                                        //2020-02-08 Start edit: gán giá trị BWCNCHob10_HasBWOrders lên server
                                        //BWCNCHob10_HasBWOrders = true;
                                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                                 //2020-02-08 End edit
                                    }

                                    break;//chỉ đẩy 1 group
                                }
                            }
                            else if (g.teethShape == orderBWHob12_GroupItems.LastOrDefault().teethShape)
                            {
                                //nếu đây là group cuối
                                //đẩy group cuối qua cho BWCNC9
                                var items = ListOrdersBWHob12.Except(ListOrdersBWCNCHob9.ToList()).Where(c => !c.TeethShape.Equals(Hob12BW_TeethShape)).GroupBy(c => c.TeethShape).LastOrDefault().ToList();
                                lst.AddRange(items);
                                //remove các items sau khi đây từ BWHob11 qua
                                if (items != null && items.Count() > 0)
                                {
                                    ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ListOrdersBWHob12.Except(items).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                    //bật flags BWCNCHob10_HasOrders
                                    //2020-02-08 Start edit: gán giá trị BWCNCHob9_HasBWOrders trên server
                                    //BWCNCHob9_HasBWOrders = true;
                                    MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(1);//true
                                                                                             //2020-02-08 End edit
                                }

                            }
                        }
                        //sắp xếp lại lst theo TeethShape, TeethQty, Diameter_d của BWCNCHob10
                        //var orderedBWCNCHob10 = SortOrders(lst, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        var orderedBWCNCHob10 = SetupSSD_BW(lst,SSDType.N1, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter);
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(orderedBWCNCHob10,Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //2020-02-08 Start edit:
                        //BWCNCHob10_NeedOrders = false;
                        MachineTypeBLL.Instance.EditFlagCheckBWOrder_CNCHob10(0);//false
                                                                                 //2020-02-08 End edit
                        if (orderedBWCNCHob10 != null && orderedBWCNCHob10.Count > 0)
                        {
                            string sourceJson = FormatHTMLOrders(ListOrdersBWHob12.ToList());
                            string newSourceJson = FormatHTMLOrders(lst);
                            mailBody.AppendLine("Đã chuyển " + orderedBWCNCHob10.Count + " orders qua máy CNC10 (D2)</br>");
                            mailBody.AppendLine("----------------------------------------------------------------------</br>");
                            mailBody.AppendLine("<strong> ORDERS BAN ĐẦU: </strong></br>");
                            mailBody.AppendLine(sourceJson + "</br>");
                            mailBody.AppendLine("<strong> ORDERS ĐÃ CHUYỂN: </strong></br>");
                            mailBody.AppendLine(newSourceJson + "</br>");
                            //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                            //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Chuyển đơn qua từ BW qua CNC (BW12)", mailBody.ToString(), "thienphuoc@spclt.com.vn");
                            //2020-02-27 End add
                        }
                    }
                    Calc_PO_and_Orders_BW();
                }


            }
        }

       
        //2020-03-12 End add

        //2019-12-16 End add

        //2019-12-19 Start add: thêm function cân bằng đơn giữa 2 máy
        private void BalancedOrdersBW(string modeCheck = "")
        {
            var t = Convert.ToInt32(DateTime.Now.ToString("HH"));
            int dayAddToCheckN1 = HolidayHelper.Instance.DayCountN1; //DateTime.Now.ToString("dddd").Equals("Saturday") ? 2 : 1;//nếu thứ 7 thì đơn N+1 là +2(vào thứ 2)
            //int status = 0; //status cập nhật bwIsSetupN1
            StringBuilder mailBody = new StringBuilder();
          
            //nếu từ 6:00 đến 23:59 => lúc này chỉ lấy các đơn N+1
            if(t >=6 && t<24)
            {
                if (ListOrdersBWHob11 != null && (ListOrdersBWHob12 != null && ListOrdersBWHob12.Count == 0))
                {
                    //get N+1 orders in BWHob11
                    var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();
                    
                    //2020-02-22 End edit
                    //kiểm tra nếu có đơn N+1 và thời gian đang nằm trong khoảng N+1 thì kiểm tra các group còn lại có đơn cùng TeethShape với máy còn lại ko
                    if (ordersBWHob11_N1 != null && ordersBWHob11_N1.Count > 0)
                    {
                        //chia đều số lượng orders N+1 
                        int size = Convert.ToInt32((1.0 * ordersBWHob11_N1.Count) / 2);
                
                        //đấy một nửa đơn là N+1 lên BWHob12 và sắp xếp theo điều kiện của BWHob12
                        if (ListOrdersBWHob12 == null)
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(); //nếu BWHob12 null => khởi tạo tránh Exception

                        var ordersForBWHob12 = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersBWHob11_N1, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter).Take(size).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode);
                                               
                        //add thêm đơn N+1 từ BW11 đưa qua                    
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ordersForBWHob12);
                        
                        //remove các order đã đẩy qua cho BWHob12
                        var result = ListOrdersBWHob11.Except(ListOrdersBWHob12).ToList();
                        //đưa các orders N+1 còn lại bên BWHob11 lên trước
                        ordersBWHob11_N1 = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList(),Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter,Hob11BW_GlobalCode);
                        if(ordersBWHob11_N1.Count > 0)
                        {
                            //lấy order cuối cùng của list order N+1 để sắp xếp các order thường theo shape, quantity và diameter
                            var last_order = ordersBWHob11_N1.LastOrDefault();
                            var ordersNotN1 = SortOrders(result.Except(ordersBWHob11_N1).ToList(), last_order.TeethShape, last_order.TeethQuantity, last_order.Diameter_d).ToList();
                            //refresh BWHob11
                            List<OrderModel> temp = new List<OrderModel>();
                            temp.AddRange(ordersBWHob11_N1);
                            temp.AddRange(ordersNotN1);
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(temp, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                        }
                        else
                        {
                            //nếu sau khi chuyển orders N+1 qua BW12 mà BW11 hết N+1 thì refresh đơn thường
                            //sắp xếp lại các orders còn lại của BWHob11
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode);
                            //refresh orders Hob11
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedOrders);
                        }
                                                                 
                        //cập nhật nội dung mail
                        mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                        mailBody.AppendLine("Cân bằng số lượng orders BW11 cho BW12, có orders N + 1(A1)</br>");
                        //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                        Calc_PO_and_Orders_BW();

                    }
                    else if (ordersBWHob11_N1 == null || ordersBWHob11_N1.Count == 0)
                    {
                        //nếu ko có đơn N+1, lấy ra các group mà ko cùng với TeethShape BWHob11 đang in (trường hợp ListOrdersBWHob11 cũng ko có đơn nào cùng TeethShape thì bỏ group đầu)
                        var group_ordersBWHob11 = ListOrdersBWHob11.GroupBy(c => c.TeethShape).Select(c => new { teethShape = c.Key, orders = c.ToList() }).ToList();
                        var ordersForBWHob12 = group_ordersBWHob11.Skip(1).ToList().Where(c=>c.teethShape.Equals(Hob12BW_TeethShape));//bỏ group đầu tiên của BWHob11
                        //kiểm tra các group có cùng shape với BW12 thì lấy
                        if (ordersForBWHob12 != null && ordersForBWHob12.Count() > 0)
                        {
                            //đẩy group có cùng shape với BW12 qua BW12
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersForBWHob12.FirstOrDefault().orders, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode);
                            //refresh BW12
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(sortedOrders);

                            //remove các orders đã đẩy qua BWHob12 và sắp xếp lại
                            var result = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ListOrdersBWHob11.Except(ListOrdersBWHob12).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode);
                            //refresh orders Hob11
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(result);
                            //cập nhật nội dung mail
                            mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                            mailBody.AppendLine("Cân bằng số lượng orders BW11 cho BW12, không có orders N + 1(A2)</br>");
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                            Calc_PO_and_Orders_BW();

                        }
                        else
                        {
                            //nếu ko có group nào cùng shape với BW12 thì lấy group cuối
                            var lastGroup = group_ordersBWHob11.Skip(1).LastOrDefault();
                            if(lastGroup != null && lastGroup.orders.Count > 0)
                            {
                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(lastGroup.orders,Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                //remove các orders đã đẩy qua BWHob12 và sắp xếp lại
                                var result = ListOrdersBWHob11.Except(ListOrdersBWHob12).ToList();
                                //refresh orders Hob11
                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                //cập nhật nội dung mail
                                mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                                mailBody.AppendLine("Cân bằng số lượng orders BW11 cho BW12, không có orders N + 1(A3)</br>");
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                                Calc_PO_and_Orders_BW();

                            }
                        }

                       
                    }
                
                }
                else if (ListOrdersBWHob12 != null && (ListOrdersBWHob11 != null && ListOrdersBWHob11.Count == 0))
                {
                    //get N+1 orders in BWHob12 (đơn N)
                    var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();

                    //2020-02-22 End edit
                    //kiểm tra nếu có đơn N+1 và thời gian đang nằm trong khoảng N+1 thì kiểm tra các group còn lại có đơn cùng TeethShape với máy còn lại ko
                    if (ordersBWHob12_N1 != null && ordersBWHob12_N1.Count > 0)
                    {
                        //chia đều số lượng orders N+1 
                        int size = Convert.ToInt32((1.0 * ordersBWHob12_N1.Count) / 2);

                        //đấy một nửa các đơn là N+1 lên BWHob11 và sắp xếp theo điều kiện của BWHob11
                        if (ListOrdersBWHob11 == null)
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(); //nếu BWHob11 null => khởi tạo tránh Exception

                        var ordersForBWHob11 = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersBWHob12_N1, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter).Take(size).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode);

                        //add thêm đơn N+1 từ BW12 đưa qua                    
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ordersForBWHob11);

                        //remove các order đã đẩy qua cho BWHob11
                        var result = ListOrdersBWHob12.Except(ListOrdersBWHob11).ToList();
                        
                        //đưa các orders N+1 còn lại bên BWHob12 lên trước
                        ordersBWHob12_N1 = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode);
                        if(ordersBWHob12_N1.Count >0)
                        {
                            //lấy order cuối cùng của list order N+1 để sắp xếp các order thường theo shape, quantity và diameter
                            var last_order = ordersBWHob12_N1.LastOrDefault();
                            var ordersNotN1 = SortOrders(result.Except(ordersBWHob12_N1).ToList(), last_order.TeethShape, last_order.TeethQuantity, last_order.Diameter_d).ToList();
                            //refresh BWHob12
                            List<OrderModel> temp = new List<OrderModel>();
                            temp.AddRange(ordersBWHob12_N1);
                            temp.AddRange(ordersNotN1);
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(temp, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                        }
                        else 
                        {
                            //nếu chia hết đơn N+1 cho BW11 thì BW12 sắp xếp thường các orders còn lại 
                            //sắp xếp lại các orders còn lại của BWHob12
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode);

                            //refresh orders Hob12
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(sortedOrders);
                        }

                        //cập nhật nội dung mail
                        mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                        mailBody.AppendLine("Cân bằng số lượng orders BW12 cho BW11, có orders N + 1(B1)</br>");
                        //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                        Calc_PO_and_Orders_BW();

                    }
                    else if (ordersBWHob12_N1 == null || ordersBWHob12_N1.Count == 0)
                    {
                        //nếu ko có đơn N+1, lấy ra các group mà ko cùng với TeethShape BWHob12 đang in (trường hợp ListOrdersBWHob12 cũng ko có đơn nào cùng TeethShape thì bỏ group đầu)
                        var group_ordersBWHob12 = ListOrdersBWHob12.GroupBy(c => c.TeethShape).Select(c => new { teethShape = c.Key, orders = c.ToList() }).ToList();
                        var ordersForBWHob11 = group_ordersBWHob12.Skip(1).ToList().Where(c => c.teethShape.Equals(Hob12BW_TeethShape));//bỏ group đầu tiên của BWHob12
                        //kiểm tra các group có cùng shape với BW11 thì lấy
                        if (ordersForBWHob11 != null && ordersForBWHob11.Count() > 0)
                        {
                            //đẩy group có cùng shape với BW12 qua BW12
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersForBWHob11.FirstOrDefault().orders, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode); 
                            //refresh BW11
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedOrders);

                            //remove các orders đã đẩy qua BWHob11 và sắp xếp lại
                            var result = ListOrdersBWHob12.Except(ListOrdersBWHob11).ToList();
                            //refresh orders Hob11
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                            //cập nhật nội dung mail
                            mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                            mailBody.AppendLine("Cân bằng số lượng orders BW12 cho BW11, không có orders N + 1(B2)</br>");
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                            Calc_PO_and_Orders_BW();

                        }
                        else
                        {
                            //nếu ko có group nào cùng shape với BW11 thì lấy group cuối
                            var lastGroup = group_ordersBWHob12.Skip(1).LastOrDefault();
                            if (lastGroup != null && lastGroup.orders.Count > 0)
                            {
                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(lastGroup.orders, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                //remove các orders đã đẩy qua BWHob11 và sắp xếp lại
                                var result = ListOrdersBWHob12.Except(ListOrdersBWHob11).ToList();
                                //refresh orders Hob12
                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                //cập nhật nội dung mail
                                mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                                mailBody.AppendLine("Cân bằng số lượng orders BW12 cho BW11, không có orders N + 1(B3)</br>");
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                                Calc_PO_and_Orders_BW();

                            }
                        }

                    }

                }
            }
            else if(t >=0 && t<6)
            {
                //nếu thời gian từ 0:00 đến 6:00 => lấy đơn N
                if (ListOrdersBWHob11 != null && (ListOrdersBWHob12 != null && ListOrdersBWHob12.Count == 0))
                {
                    //get orders in BWHob11 (đơn N)
                    var ordersBWHob11_N = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();

                    //2020-02-22 End edit
                    //kiểm tra nếu có đơn N và thời gian đang nằm trong khoảng N+1 thì kiểm tra các group còn lại có đơn cùng TeethShape với máy còn lại ko
                    if (ordersBWHob11_N != null && ordersBWHob11_N.Count > 0)
                    {
                        //chia đều số lượng orders N 
                        int size = Convert.ToInt32((1.0 * ordersBWHob11_N.Count) / 2);

                        //đấy một nửa các đơn là N lên BWHob12 và sắp xếp theo điều kiện của BWHob12
                        if (ListOrdersBWHob12 == null)
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(); //nếu BWHob12 null => khởi tạo tránh Exception

                        var ordersForBWHob12 = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersBWHob11_N, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter).Take(size).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode);

                        //add thêm đơn N từ BW11 đưa qua                    
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(ordersForBWHob12);

                        //remove các order đã đẩy qua cho BWHob12
                        var result = ListOrdersBWHob11.Except(ListOrdersBWHob12).ToList();
                        //đưa các orders N còn lại bên BWHob11 lên trước
                        ordersBWHob11_N = SortOrders(result.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter).ToList();
                        if(ordersBWHob11_N.Count > 0)
                        {
                            //lấy order cuối cùng của list order N để sắp xếp các order thường theo shape, quantity và diameter
                            var last_order = ordersBWHob11_N.LastOrDefault();
                            var ordersNotN = SortOrders(result.Except(ordersBWHob11_N).ToList(), last_order.TeethShape, last_order.TeethQuantity, last_order.Diameter_d).ToList();
                            //refresh BWHob11
                            List<OrderModel> temp = new List<OrderModel>();
                            temp.AddRange(ordersBWHob11_N);
                            temp.AddRange(ordersNotN);
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(temp, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                        }
                        else
                        {
                            //nếu sau khi chuyển orders N qua BW12 mà BW11 hết N thì refresh đơn thường
                            //sắp xếp lại các orders còn lại của BWHob11
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode);
                            //refresh orders Hob11
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedOrders);
                        }
                       

                        //cập nhật nội dung mail
                        mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                        mailBody.AppendLine("Cân bằng số lượng orders BW11 cho BW12, có orders N (C1)</br>");
                        //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                        Calc_PO_and_Orders_BW();

                    }
                    else if (ordersBWHob11_N == null || ordersBWHob11_N.Count == 0)
                    {
                        //nếu ko có đơn N, lấy ra các group mà ko cùng với TeethShape BWHob11 đang in (trường hợp ListOrdersBWHob11 cũng ko có đơn nào cùng TeethShape thì bỏ group đầu)
                        var group_ordersBWHob11 = ListOrdersBWHob11.GroupBy(c => c.TeethShape).Select(c => new { teethShape = c.Key, orders = c.ToList() }).ToList();
                        var ordersForBWHob12 = group_ordersBWHob11.Skip(1).ToList().Where(c => c.teethShape.Equals(Hob12BW_TeethShape));//bỏ group đầu tiên của BWHob11
                        //kiểm tra các group có cùng shape với BW12 thì lấy
                        if (ordersForBWHob12 != null && ordersForBWHob12.Count() > 0)
                        {
                            //đẩy group có cùng shape với BW12 qua BW12
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersForBWHob12.FirstOrDefault().orders, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode);
                            //refresh BW12
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(sortedOrders);

                            //remove các orders đã đẩy qua BWHob12 và sắp xếp lại
                            var result = ListOrdersBWHob11.Except(ListOrdersBWHob12).ToList();
                            //refresh orders Hob11
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                            //cập nhật nội dung mail
                            mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                            mailBody.AppendLine("Cân bằng số lượng orders BW11 cho BW12, không có orders N (C2)</br>");
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                            Calc_PO_and_Orders_BW();

                        }
                        else
                        {
                            //nếu ko có group nào cùng shape với BW12 thì lấy group cuối
                            var lastGroup = group_ordersBWHob11.Skip(1).LastOrDefault();
                            if (lastGroup != null && lastGroup.orders.Count > 0)
                            {
                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(lastGroup.orders, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                //remove các orders đã đẩy qua BWHob12 và sắp xếp lại
                                var result = ListOrdersBWHob11.Except(ListOrdersBWHob12).ToList();
                                //refresh orders Hob11
                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                //cập nhật nội dung mail
                                mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                                mailBody.AppendLine("Cân bằng số lượng orders BW11 cho BW12, không có orders N (C3)</br>");
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                                Calc_PO_and_Orders_BW();

                            }
                        }


                    }

                }
                else if (ListOrdersBWHob12 != null && (ListOrdersBWHob11 != null && ListOrdersBWHob11.Count == 0))
                {
                    //get N orders in BWHob12 (đơn N)
                    var ordersBWHob12_N = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();

                    //2020-02-22 End edit
                    //kiểm tra nếu có đơn và thời gian đang nằm trong khoảng thì kiểm tra các group còn lại có đơn cùng TeethShape với máy còn lại ko
                    if (ordersBWHob12_N != null && ordersBWHob12_N.Count > 0)
                    {
                        //chia đều số lượng orders N 
                        int size = Convert.ToInt32((1.0 * ordersBWHob12_N.Count) / 2);

                        //đấy hết các đơn là N lên BWHob11 và sắp xếp theo điều kiện của BWHob11
                        if (ListOrdersBWHob11 == null)
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(); //nếu BWHob11 null => khởi tạo tránh Exception

                        var ordersForBWHob11 = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersBWHob12_N, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter).Take(size).ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode);

                        //add thêm đơn N từ BW12 đưa qua                    
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(ordersForBWHob11);

                        //remove các order đã đẩy qua cho BWHob11
                        var result = ListOrdersBWHob12.Except(ListOrdersBWHob11).ToList();

                        //đưa các orders N còn lại bên BWHob12 lên trước
                        ordersBWHob12_N = SortOrders(result.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList(), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter).ToList();
                        if(ordersBWHob12_N.Count >0)
                        {
                            //lấy order cuối cùng của list order N để sắp xếp các order thường theo shape, quantity và diameter
                            var last_order = ordersBWHob12_N.LastOrDefault();
                            var ordersNotN = SortOrders(result.Except(ordersBWHob12_N).ToList(), last_order.TeethShape, last_order.TeethQuantity, last_order.Diameter_d).ToList();
                            //refresh BWHob12
                            List<OrderModel> temp = new List<OrderModel>();
                            temp.AddRange(ordersBWHob12_N);
                            temp.AddRange(ordersNotN);
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(temp, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                        }
                        else
                        {
                            //sắp xếp lại các orders còn lại của BWHob12
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode);

                            //refresh orders Hob12
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(sortedOrders);

                        }

                        //cập nhật nội dung mail
                        mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                        mailBody.AppendLine("Cân bằng số lượng orders BW12 cho BW11, có orders N (D1)</br>");
                        //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                        Calc_PO_and_Orders_BW();

                    }
                    else if (ordersBWHob12_N == null || ordersBWHob12_N.Count == 0)
                    {
                        //nếu ko có đơn N, lấy ra các group mà ko cùng với TeethShape BWHob12 đang in (trường hợp ListOrdersBWHob12 cũng ko có đơn nào cùng TeethShape thì bỏ group đầu)
                        var group_ordersBWHob12 = ListOrdersBWHob12.GroupBy(c => c.TeethShape).Select(c => new { teethShape = c.Key, orders = c.ToList() }).ToList();
                        var ordersForBWHob11 = group_ordersBWHob12.Skip(1).ToList().Where(c => c.teethShape.Equals(Hob12BW_TeethShape));//bỏ group đầu tiên của BWHob12
                        //kiểm tra các group có cùng shape với BW11 thì lấy
                        if (ordersForBWHob11 != null && ordersForBWHob11.Count() > 0)
                        {
                            //đẩy group có cùng shape với BW12 qua BW12
                            var sortedOrders = BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(ordersForBWHob11.FirstOrDefault().orders, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode);
                            //refresh BW11
                            ListOrdersBWHob11 = new ObservableCollection<OrderModel>(sortedOrders);

                            //remove các orders đã đẩy qua BWHob11 và sắp xếp lại
                            var result = ListOrdersBWHob12.Except(ListOrdersBWHob11).ToList();
                            //refresh orders Hob11
                            ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                            //cập nhật nội dung mail
                            mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                            mailBody.AppendLine("Cân bằng số lượng orders BW12 cho BW11, không có orders N (D2)</br>");
                            //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                            Calc_PO_and_Orders_BW();

                        }
                        else
                        {
                            //nếu ko có group nào cùng shape với BW11 thì lấy group cuối
                            var lastGroup = group_ordersBWHob12.Skip(1).LastOrDefault();
                            if (lastGroup != null && lastGroup.orders.Count > 0)
                            {
                                ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(lastGroup.orders, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                                //remove các orders đã đẩy qua BWHob11 và sắp xếp lại
                                var result = ListOrdersBWHob12.Except(ListOrdersBWHob11).ToList();
                                //refresh orders Hob12
                                ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(result, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                                //cập nhật nội dung mail
                                mailBody.AppendLine("Thời gian: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "</br>");
                                mailBody.AppendLine("Cân bằng số lượng orders BW12 cho BW11, không có orders N (D3)</br>");
                                //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - " + modeCheck, mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                                Calc_PO_and_Orders_BW();

                            }
                        }


                    }

                }

            }

          
            //cập nhật bwIsSetupN1
            //try
            //{
            //    OrderBLL.Instance.Update_bwIsSetup(status);
            //}
            //catch (Exception ex)
            //{

            //    throw new Exception("Cập nhật trạng thái orders N+1 thất bại\n"+ex.Message);
            //}

        }

        private void BalancedOrdersBW_NormalMode()
        {
            StringBuilder mailBody = new StringBuilder();
            int dayAddToCheckN1 = HolidayHelper.Instance.DayCountN1; //DateTime.Now.ToString("dddd").Equals("Saturday") ? 2 : 1;//nếu thứ 7 thì đơn N+1 là +2(vào thứ 2)
            string currentDayName = DateTime.Now.ToString("dddd");
            if (ListOrdersBWHob11 != null && (ListOrdersBWHob12 != null && ListOrdersBWHob12.Count == 0))
            {
                //get N+1 orders in BWHob11 (đơn N)
                var ordersBWHob11_N1 = ListOrdersBWHob11.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();

                //2020-02-22 Start add: lấy các đơn N+1 today + 1
                var ordersBWHob11_TodayAdd1 = ListOrdersBWHob11.Except(ordersBWHob11_N1).Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();
                //2020-02-22 End add

                //2020-02-22 Start edit: tách lấy các group có đơn ko phải N+1 và Today+1
                //var orderBWHob11_group = ListOrdersBWHob11.Except(ordersBWHob11_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, orders = g.ToList() }).ToList();
                var orderBWHob11_group = ListOrdersBWHob11.Except(ordersBWHob11_N1).Except(ordersBWHob11_TodayAdd1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, orders = g.ToList() }).ToList();
                //2020-02-22 End edit
                //kiểm tra nếu có đơn N+1 và ngày hiện tại là thứ 7 => vẫn đẩy đơn N+1 lên 
                 if ((ordersBWHob11_N1 != null && ordersBWHob11_N1.Count > 0) || (ordersBWHob11_TodayAdd1 != null && ordersBWHob11_TodayAdd1.Count > 0))
                 {
                    ////vì thứ 7 ko có ca đêm nên ko cần kiểm tra hàng N
                    //if (currentDayName.Equals("Saturday"))
                    //{
                    //    //sort order N+1
                    //    var sortedN1 = SortOrders(ordersBWHob11_TodayAdd1, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter);
                    //    //lấy ra orderN1 cuối để xếp teeth cho các order thường còn lại theo orderN1 cuối
                    //    var lastN1 = sortedN1.LastOrDefault();
                    //    var sortedNotN1 = SortOrders(ListOrdersBWHob11.Except(sortedN1).ToList(), lastN1.TeethShape, lastN1.TeethQuantity, lastN1.Diameter_d);
                    //    List<OrderModel> temp = new List<OrderModel>();
                    //    temp.AddRange(sortedN1);
                    //    temp.AddRange(sortedNotN1);
                    //    //refresh ListOrdersBWHob11
                    //    ListOrdersBWHob11 = new ObservableCollection<OrderModel>(temp);
                    //}
                 }
                 else
                 {
                    //nếu ko có đơn N+1, lấy ra các group mà ko cùng với TeethShape BWHob11 đang in (trường hợp ListOrdersBWHob11 cũng ko có đơn nào cùng TeethShape thì bỏ group đầu)
                    var group_ordersBWHob11 = ListOrdersBWHob11.GroupBy(c => c.TeethShape).Select(c => new { teethShape = c.Key, orders = c.ToList() }).ToList();
                    var ordersForBWHob12 = group_ordersBWHob11.Skip(1).ToList();//bỏ group đầu tiên của BWHob11

                    //đấy hết các đơn không phải N+1 lên BWHob12 và sắp xếp theo điều kiện của BWHob12
                    if (ListOrdersBWHob12 == null)
                        ListOrdersBWHob12 = new ObservableCollection<OrderModel>(); //nếu BWHob12 null => khởi tạo tránh Exception

                    var currentOrdersInBWHob12 = ListOrdersBWHob12.ToList();
                    //add thêm đơn từ BW11 đưa qua

                    ordersForBWHob12.ForEach(c => currentOrdersInBWHob12.AddRange(c.orders));
                    ListOrdersBWHob12 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(currentOrdersInBWHob12, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter), Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter, Hob12BW_GlobalCode));
                    //remove các order đã đẩy qua cho BWHob12
                    var result = ListOrdersBWHob11.Except(ListOrdersBWHob12);
                    //refresh BWHob11
                    ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(result.ToList(), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                    mailBody.AppendLine("Số orders máy BW11 hiện tại: "+ListOrdersBWHob11.Count);
                    mailBody.AppendLine("Số orders máy BW12 hiện tại: " + ListOrdersBWHob12.Count);
                    //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                    //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                    //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - NormalMode", mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                    //2020-02-27 End add
                }

            }

            //nếu BWHob12 khác null và BWHob11 hết đơn
            if (ListOrdersBWHob12 != null && (ListOrdersBWHob11 != null && ListOrdersBWHob11.Count == 0))
            {

                //get N+1 orders in BWHob12 (đơn N)
                var ordersBWHob12_N1 = ListOrdersBWHob12.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();

                //2020-02-22 Start add: lấy các đơn today + 1
                var ordersBWHob12_TodayAdd1 = ListOrdersBWHob12.Except(ordersBWHob12_N1).Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayAddToCheckN1).Date).ToList();
                //2020-02-22 End add

                //2020-02-22 Start edit: tách lấy các group có đơn ko phải N+1 và today+1
                //var orderBWHob12_group = ListOrdersBWHob12.Except(ordersBWHob12_N1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, orders = g.ToList() }).ToList();
                var orderBWHob12_group = ListOrdersBWHob12.Except(ordersBWHob12_N1).Except(ordersBWHob12_TodayAdd1).GroupBy(c => c.TeethShape).Select(g => new { teethShape = g.Key, orders = g.ToList() }).ToList();
                //2020-02-22 End edit
                //kiểm tra nếu có đơn N+1 và ngày hiện tại là thứ 7 => vẫn đẩy đơn N+1 lên 
                if ((ordersBWHob12_N1 != null && ordersBWHob12_TodayAdd1.Count > 0) || (ordersBWHob12_TodayAdd1 != null && ordersBWHob12_TodayAdd1.Count > 0))
                {
                    ////vì thứ 7 ko có ca đêm nên ko cần kiểm tra hàng N
                    //if (currentDayName.Equals("Saturday"))
                    //{
                    //    //sort order N+1
                    //    var sortedN1 = SortOrders(ordersBWHob12_TodayAdd1, Hob12BW_TeethShape, Hob12BW_TeethQty, Hob12BW_Diameter);
                    //    //lấy ra orderN1 cuối để xếp teeth cho các order thường còn lại theo orderN1 cuối
                    //    var lastN1 = sortedN1.LastOrDefault();
                    //    var sortedNotN1 = SortOrders(ListOrdersBWHob12.Except(sortedN1).ToList(), lastN1.TeethShape, lastN1.TeethQuantity, lastN1.Diameter_d);
                    //    List<OrderModel> temp = new List<OrderModel>();
                    //    temp.AddRange(sortedN1);
                    //    temp.AddRange(sortedNotN1);
                    //    //refresh ListOrdersBWHob12
                    //    ListOrdersBWHob12 = new ObservableCollection<OrderModel>(temp);
                    //}
                }
                else
                {
                    //nếu ko có đơn N+1, lấy ra các group mà ko cùng với TeethShape BWHob12 đang in (trường hợp ListOrdersBWHob12 cũng ko có đơn nào cùng TeethShape thì bỏ group đầu)
                    var group_ordersBWHob12 = ListOrdersBWHob12.GroupBy(c => c.TeethShape).Select(c => new { teethShape = c.Key, orders = c.ToList() }).ToList();
                    var ordersForBWHob11 = group_ordersBWHob12.Skip(1).ToList();//bỏ group đầu tiên của BWHob12

                    //đấy hết các đơn không phải N+1 lên BWHob11 và sắp xếp theo điều kiện của BWHob11
                    if (ListOrdersBWHob11 == null)
                        ListOrdersBWHob11 = new ObservableCollection<OrderModel>(); //nếu BWHob11 null => khởi tạo tránh Exception

                    var currentOrdersInBWHob11 = ListOrdersBWHob11.ToList();
                    //add thêm đơn từ BW11 đưa qua

                    ordersForBWHob11.ForEach(c => currentOrdersInBWHob11.AddRange(c.orders));
                    ListOrdersBWHob11 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(SortOrders(currentOrdersInBWHob11, Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter), Hob11BW_TeethShape, Hob11BW_TeethQty, Hob11BW_Diameter, Hob11BW_GlobalCode));
                    //remove các order đã đẩy qua cho BWHob11
                    var result = ListOrdersBWHob12.Except(ListOrdersBWHob11);
                    //refresh BWHob12
                    ListOrdersBWHob12 = new ObservableCollection<OrderModel>(result);
                    mailBody.AppendLine("Số orders máy BW11 hiện tại: " + ListOrdersBWHob11.Count);
                    mailBody.AppendLine("Số orders máy BW12 hiện tại: " + ListOrdersBWHob12.Count);
                    //2020-02-27 Start add: thêm mail thông báo việc sắp xếp cân bằng orders máy BW
                    //"honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "[AutomationMailing] Pulley's PO automation tool", mailBody, "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn"
                    //SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "honghanh@spclt.com.vn", "Thông báo cân bằng orders các máy BW - NormalMode", mailBody.ToString(), "be.hua@spclt.com.vn;thienphuoc@spclt.com.vn;thuc.nguyen@spclt.com.vn");
                    //2020-02-27 End add
                }
            }

            //try
            //{
            //    OrderBLL.Instance.Update_bwIsSetup(0); //cập nhật status bwIsSetupN1 = 0
            //}
            //catch (Exception ex)
            //{

            //    throw new Exception("Cập nhật trạng thái Orders N+1 = 0 thất bại\n"+ex.Message);
            //}

            Calc_PO_and_Orders_BW();
        }

        private void BalancedOrdersCNC(List<OrderModel> remainList = null)
        {
            //điều kiện đầu tiên check trong lúc lấy đơn CNC về
            if(remainList != null)
            {
                var orders = SortOrders_CNC(remainList);
                var groupOrders = orders.GroupBy(c => c.Material_text1).Select(c => new { k = c.Key, items = c.ToList() });//Material_text1 = D_Mat
                var groupCountOrders = groupOrders.Count();
                var numOfGroupToTake = Math.Ceiling((groupCountOrders * 1.0) / 2);


                if (ListOrdersCNCHob9 == null)
                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>();
                else
                {
                    List<OrderModel> currentListOrderCNCHob9 = ListOrdersCNCHob9.ToList();
                    //add more items from groupOrders
                    var g = groupOrders.Take((int)numOfGroupToTake).ToList();
                    g.ForEach(c => currentListOrderCNCHob9.AddRange(c.items));

                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(currentListOrderCNCHob9);
                }

                if (ListOrdersCNCHob10 == null)
                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>();
                else
                {
                    List<OrderModel> currentListOrderCNCHob10 = ListOrdersCNCHob10.ToList();
                    var remainOrdersFromListOrderCNCHob9 = orders.Except(ListOrdersCNCHob9);
                    currentListOrderCNCHob10.AddRange(remainOrdersFromListOrderCNCHob9);
                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(currentListOrderCNCHob10);
                }
            }
            else 
            {
                //điều kiện này được check khi timerCNC chạy
                //nếu CNCHob9 có đơn mà CNCHob10 vẫn ko có đơn
                //chia bớt từ CNCHob9 qua
                if (ListOrdersCNCHob9 != null && ListOrdersCNCHob9.Count >= 2)
                {
                    if (ListOrdersCNCHob10 == null)
                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>();
                    else
                    {
                        if(ListOrdersCNCHob10.Count == 0)
                        {
                            var count = ListOrdersCNCHob9.Count;
                            var numOfOrdersToTake = Math.Ceiling((count) * 1.0 / 2);
                            var beginItems = ListOrdersCNCHob9.Take((int)numOfOrdersToTake);
                            ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ListOrdersCNCHob9.Except(beginItems));
                            //refresh list CNCHob9 và remove các orders đã đẩy qua CNCHob10
                            ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ListOrdersCNCHob9.Except(ListOrdersCNCHob10));
                        }
                    }                
                }

                if (ListOrdersCNCHob10 != null && ListOrdersCNCHob10.Count >= 2)
                {
                    if (ListOrdersCNCHob9 == null)
                        ListOrdersCNCHob9 = new ObservableCollection<OrderModel>();
                    else
                    {
                        if(ListOrdersCNCHob9.Count == 0)
                        {
                            var count = ListOrdersCNCHob10.Count;
                            var numOfOrdersToTake = Math.Ceiling((count) * 1.0 / 2);
                            var beginItems = ListOrdersCNCHob10.Take((int)numOfOrdersToTake);
                            ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ListOrdersCNCHob10.Except(beginItems));
                            //refresh list CNCHob10 và remove các orders đã đẩy qua CNCHob9
                            ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ListOrdersCNCHob10.Except(ListOrdersCNCHob9));
                        }
                    }
                   
                }
            }

            Calc_PO_and_Orders_CNC();

        }
        //2020-02-08 Start add: chỉnh lại chức năng cân bằng order giữa 2 máy CNC bên màn hình CNC     
        private void BalancedOrdersCNC2()
        {
            //từ CNCHob9 qua CNCHob10: nếu ListOrdersCNCHob10 hết đơn
            if (ListOrdersCNCHob10 != null && ListOrdersCNCHob10.Count == 0)
            {
                if (ListOrdersCNCHob9 != null)
                {
                    //group các item trên list theo D_mat
                    var groupOrders = ListOrdersCNCHob9.GroupBy(c => c.Material_text1).Select(c => new { k = c.Key, items = c.ToList() });//Material_text1 = D_Mat                
                    var count = groupOrders.Count();
                    //nếu nhiều hơn 1 group thì đẩy group cuối cùng các group còn lại qua
                    if (count > 1)
                    {
                        var ordersFor_CNCHob10 = groupOrders.Skip(1).LastOrDefault().items.ToList();//ko sort orders vì CNCHob9 đã sort
                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ordersFor_CNCHob10);
                        //remove các orders đã chuyển qua CNCHob10
                        var remainOrdersForCNCHob9 = ListOrdersCNCHob9.Except(ListOrdersCNCHob10);
                        ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(remainOrdersForCNCHob9);
                    }
                }
            }


            //từ CNCHob10 qua CNCHob9: nếu ListOrdersCNCHob9 hết đơn
            if (ListOrdersCNCHob9 != null && ListOrdersCNCHob9.Count == 0)
            {
                if (ListOrdersCNCHob10 != null)
                {
                    //group các item trên list theo D_mat
                    var groupOrders = ListOrdersCNCHob9.GroupBy(c => c.Material_text1).Select(c => new { k = c.Key, items = c.ToList() });//Material_text1 = D_Mat                
                    var count = groupOrders.Count();
                    //nếu nhiều hơn 1 group thì đẩy group cuối cùng các group còn lại qua
                    if (count > 1)
                    {
                        var ordersFor_CNCHob9 = groupOrders.Skip(1).LastOrDefault().items.ToList();//ko sort orders vì CNCHob10 đã sort
                        ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ordersFor_CNCHob9);
                        //remove các orders đã chuyển qua CNCHob9
                        var remainOrdersForCNCHob10 = ListOrdersCNCHob10.Except(ListOrdersCNCHob9);
                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(remainOrdersForCNCHob10);
                    }
                }
            }

            Calc_PO_and_Orders_CNC();

        }
        //2020-02-08 End add

        //2020-02-08 Start add: chỉnh lại chức năng cân bằng order giữa 2 máy CNC bên màn hình BW     
        private void BalancedOrdersBWCNC()
        {
            //từ BWCNCHob9 qua BWCNCHob10: nếu ListOrdersBWCNCHob10 hết đơn
            if (ListOrdersBWCNCHob10 != null && ListOrdersBWCNCHob10.Count == 0)
            {
                if (ListOrdersBWCNCHob9 != null)
                {
                    //group các item trên list theo D_mat
                    var groupOrders = ListOrdersBWCNCHob9.GroupBy(c => c.Material_text1).Select(c => new { k = c.Key, items = c.ToList() });//Material_text1 = D_Mat                
                    var count = groupOrders.Count();
                    //nếu nhiều hơn 1 group thì đẩy group cuối cùng các group còn lại qua
                    if (count > 1)
                    {
                        var ordersFor_BWCNCHob10 = groupOrders.Skip(1).LastOrDefault().items.ToList();//ko sort orders vì BWCNCHob9 đã sort
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ordersFor_BWCNCHob10, Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                        //remove các orders đã chuyển qua BWCNCHob10
                        var remainOrdersForBWCNCHob9 = ListOrdersBWCNCHob9.Except(ListOrdersBWCNCHob10);
                        ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(remainOrdersForBWCNCHob9.ToList(),Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));
                    }
                }
            }


            //từ BWCNCHob10 qua BWCNCHob9: nếu ListOrdersBWCNCHob9 hết đơn
            if (ListOrdersBWCNCHob9 != null && ListOrdersBWCNCHob9.Count == 0)
            {
                if (ListOrdersBWCNCHob10 != null)
                {
                    //group các item trên list theo D_mat
                    var groupOrders = ListOrdersBWCNCHob9.GroupBy(c => c.Material_text1).Select(c => new { k = c.Key, items = c.ToList() });//Material_text1 = D_Mat                
                    var count = groupOrders.Count();
                    //nếu nhiều hơn 1 group thì đẩy group cuối cùng các group còn lại qua
                    if (count > 1)
                    {
                        var ordersFor_BWCNCHob9 = groupOrders.Skip(1).LastOrDefault().items.ToList();//ko sort orders vì BWCNCHob10 đã sort
                        ListOrdersBWCNCHob9 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(ordersFor_BWCNCHob9, Hob9BWCNC_TeethShape, Hob9BWCNC_TeethQty, Hob9BWCNC_Diameter, Hob9BWCNC_GlobalCode));
                        //remove các orders đã chuyển qua BWCNCHob9
                        var remainOrdersForBWCNCHob10 = ListOrdersBWCNCHob10.Except(ListOrdersBWCNCHob9);
                        ListOrdersBWCNCHob10 = new ObservableCollection<OrderModel>(BWCNCViewModel_Sub.Instance.SortByWithGlobalCode(remainOrdersForBWCNCHob10.ToList(), Hob10BWCNC_TeethShape, Hob10BWCNC_TeethQty, Hob10BWCNC_Diameter, Hob10BWCNC_GlobalCode));
                    }
                }
            }

            Calc_PO_and_Orders_BW();

        }
        //2020-02-08 End add
        //2019-12-19 End add

        //2020-02-21 Start add: 
        //sub functions to calculate SPH 
        //mode 1
        private Tuple<bool,double,int,string> CheckSPH_Mode1_BW(int hour,List<OrderModel> lstOrders)
        {
            int dayAddToCheckN1 = HolidayHelper.Instance.DayCountN1; //DateTime.Now.ToString("dddd").Equals("Saturday") ? 2 : 1;//nếu thứ 7 thì ngày check N+1 phải +2
            int sum_totalInstruction_N1 = 0;
            double remainPT_PulleyLineB = 0;
            bool result = false;
            string jsonStr = "";
            try
            {
              
                //2020-03-09 Start edit: 

                //lấy sum các giá trị của line pulley BW
                remainPT_PulleyLineB = OrderBLL.Instance.GetSumPulley_BW(hour);

                //lấy các ngày lễ của năm hiện tại
                List<DateTime> lstHolidays = OrderBLL.Instance.GetHolidaysInYear();
               
                if (hour == 0)//nếu hour = 24 -> đơn N+1 sẽ thành N
                {
                    var tempList = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.ToString("yyyyMMdd")).ToList();
                    if (tempList != null && tempList.Count > 0)
                    {
                        sum_totalInstruction_N1 = tempList.Sum(c=>Convert.ToInt32(c.Number_of_Available_Instructions));
                        jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(tempList.Select(c => new { c.Manufa_Instruction_No, c.Factory_Ship_Date }).ToList());
                    }
                }else
                {
                    var tempList = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.AddDays(dayAddToCheckN1).ToString("yyyyMMdd")).ToList();
                    if (tempList != null && tempList.Count > 0)
                    {
                        sum_totalInstruction_N1 = tempList.Sum(c=> Convert.ToInt32(c.Number_of_Available_Instructions));
                        jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(tempList.Select(c => new {SoPO =  c.Manufa_Instruction_No, NgayXuat = c.Factory_Ship_Date }).ToList());
                    }
                }
                //nếu thứ 7 thì chỉ check 0 < N+1
                result = DateTime.Now.ToString("dddd").Equals("Saturday") ? (0 < sum_totalInstruction_N1) :  (remainPT_PulleyLineB <= sum_totalInstruction_N1);
                
            }
            catch (Exception ex)
            {

                throw ex;
            }
            
            return new Tuple<bool, double, int, string>(result,remainPT_PulleyLineB,sum_totalInstruction_N1, jsonStr);
        }

        //mode 2
        private Tuple<bool,double,int,string> CheckSPH_Mode2_BW(int hour, List<OrderModel> lstOrders)
        {
            int sum_totalInstruction_N = 0;
            //int sum_totalInstruction_N1 = 0;
         
            bool result = false;
            string jsonStr = "";
            //2020-03-09 Start edit: 

            //lấy sum các giá trị của line pulley BW
            double remainPT_PulleyLineB = OrderBLL.Instance.GetSumPulley_BW(hour);

            //lấy các ngày lễ của năm hiện tại
            List<DateTime> lstHolidays = OrderBLL.Instance.GetHolidaysInYear();

            //2020-03-09 End edit
            //2020-03-10 Start lock: khóa chỉ tính orders thuộc ngày hiện tại N
            ////get N, N+1 (not yet process qty)
            //var tempList = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.ToString("yyyyMMdd")).ToList();
            //if (tempList != null && tempList.Count > 0)
            //{
            //    sum_totalInstruction_N = tempList.Count;
            //    var tempList2 = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.AddDays(1).ToString("yyyyMMdd")).ToList();
            //    sum_totalInstruction_N1 = (tempList2 != null) ? tempList2.Count : 0;
            //}
            //else
            //{
            //    //get N+1
            //    var tempList2 = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.AddDays(1).ToString("yyyyMMdd")).ToList();
            //    sum_totalInstruction_N1 = (tempList2 != null) ? tempList2.Count : 0;
            //}
            //update remainPT máy BW, số đơn đang có N+1 
            //BW_RemainPT = remainPT_PulleyLineB;
            //BW_Total_N1_Orders = (sum_totalInstruction_N + sum_totalInstruction_N1);
            //2020-03-10 End lock
            var tempList = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.ToString("yyyyMMdd")).ToList();
            if(tempList != null && tempList.Count > 0)
            {
                sum_totalInstruction_N = tempList.Sum(c=>Convert.ToInt32(c.Number_of_Available_Instructions));
                jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(tempList.Select(c => new { SoPO = c.Manufa_Instruction_No, NgayXuat = c.Factory_Ship_Date }).ToList());
            }
            result = 0 < sum_totalInstruction_N;
            return new Tuple<bool, double, int, string>(result,remainPT_PulleyLineB,sum_totalInstruction_N,jsonStr);            
        }

        //function get workshift
        private string GetWorkShift(int hour)
        {
            //from 6:00 AM to 18:00 PM -> day shift
            return (hour >= 6 || hour < 18) ? "Day" : "Night";
            
        }

        private void ApplySSDCheck_BW(List<OrderModel> lstOrders)
        {
            var hour = Convert.ToInt32(DateTime.Now.ToString("HH"));
            //from 6:00 AM today to 2:00 AM tomorrow
            int[] hours_mode1 = { 6,7,8, 9,10,11, 12,13,14, 15,16,17, 18,19,20, 21,22,23, 0};
            int[] hours_mode2 = { 2,3,4,5};
            int[] hourSendMail_mode1 = { 6,9,12,15,18,21,0 };
            int status = 0;
            try
            {
                //status = OrderBLL.Instance.Get_bwIsSetupN1();
                //check NYP mode1
                if(hours_mode1.Contains(hour))
                {
                    var resTup = CheckSPH_Mode1_BW(hour, lstOrders);
                    //check SSD => true => cập nhật isBWCheckN1 true
                    if (resTup.Item1)
                    {
                        status = 1;
                        SSDModeCheck_BW = "SSDMode";
                        BW_RemainPT = resTup.Item2;
                        BW_Total_N1_Orders = resTup.Item3;
                    }
                    else {
                        status = 0;
                        SSDModeCheck_BW = "Min Setup Mode";
                        BW_RemainPT = resTup.Item2;
                        BW_Total_N1_Orders = resTup.Item3;
                    }
                    if (isSendMailSSDModeCheck_BW == false)
                    {
                        if (hourSendMail_mode1.Contains(hour))
                        {
                            StringBuilder builder = new StringBuilder();
                            builder.AppendLine("Mode check: " + SSDModeCheck_BW + "</br>");
                            builder.AppendLine("RemainPT x SPH: " + resTup.Item2 + "</br>");
                            builder.AppendLine("N+1 Qty: " + resTup.Item3 + "</br>");
                            builder.AppendLine("PO List:</br>");
                            builder.AppendLine("------------------------------</br>");
                            builder.AppendLine(resTup.Item4);
                            var res = SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "thienphuoc@spclt.com.vn", "Thông báo SSDMode check BW(1)", builder.ToString(), "");
                            isSendMailSSDModeCheck_BW = res != 0 ? true : false;
                        }
                    }
                }
                else if(hours_mode2.Contains(hour)) //check NYP mode2
                {
                    var resTup = CheckSPH_Mode2_BW(hour, lstOrders);
                    //check SSD => true => cập nhật isBWCheckN1 true
                    if (resTup.Item1)
                    {
                        status = 1;
                        SSDModeCheck_BW = "SSDMode";
                        BW_RemainPT = resTup.Item2;
                        BW_Total_N1_Orders = resTup.Item3;
                    }
                    else
                    {
                        status = 0;
                        SSDModeCheck_BW = "Min Setup Mode";
                        BW_RemainPT = resTup.Item2;
                        BW_Total_N1_Orders = resTup.Item3;
                    }
                    if (isSendMailSSDModeCheck_BW == false)
                    {                        
                        StringBuilder builder = new StringBuilder();
                        builder.AppendLine("Mode check: " + SSDModeCheck_BW + "</br>");
                        builder.AppendLine("RemainPT x SPH: " + resTup.Item2 + "</br>");
                        builder.AppendLine("N+1 Qty: " + resTup.Item3 + "</br>");
                        builder.AppendLine("PO List:</br>");
                        builder.AppendLine("------------------------------</br>");
                        builder.AppendLine(resTup.Item4);
                        var res = SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "thienphuoc@spclt.com.vn", "Thông báo SSDMode check BW(2)", builder.ToString(), "");
                        isSendMailSSDModeCheck_BW = res != 0 ? true : false;
                    }
                }
              

                OrderBLL.Instance.Update_bwIsSetup(status);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        //2020-03-12 Start add: thêm 3 functions check SSD cho CNC
        //check SSD CNC
        //mode 1
        private Tuple<bool, double, int, string> CheckSPH_Mode1_CNC(int hour, List<OrderModel> lstOrders)
        {

            //formular: SPH*Remain PT <= N+1 NYP Qty   
            int dayAddToCheckN1 = HolidayHelper.Instance.DayCountN1; //DateTime.Now.ToString("dddd").Equals("Saturday") ? 2 : 1;//nếu thứ 7 thì ngày check N+1 phải +2
            int sum_totalInstruction_N1 = 0;
            double remainPT_PulleyLineA = 0; //line A cho máy CNC
            bool result = false;
            string jsonStr = "";
            try
            {
                //2020-03-09 Start edit: 

                //lấy sum các giá trị của line pulley CNC
                remainPT_PulleyLineA = OrderBLL.Instance.GetSumPulley_CNC(hour);

                //lấy các ngày lễ của năm hiện tại
                List<DateTime> lstHolidays = OrderBLL.Instance.GetHolidaysInYear();
               
                if (hour == 0)//nếu hour = 24 -> đơn N+1 sẽ thành N
                {
                    var tempList = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.ToString("yyyyMMdd")).ToList();
                    if (tempList != null && tempList.Count > 0)
                    {
                        sum_totalInstruction_N1 = tempList.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions));
                        jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(tempList.Select(c => new { c.Manufa_Instruction_No, c.Factory_Ship_Date }).ToList());
                    }
                }
                else
                {
                    var tempList = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.AddDays(dayAddToCheckN1).ToString("yyyyMMdd")).ToList();
                    if (tempList != null && tempList.Count > 0)
                    {
                        sum_totalInstruction_N1 = tempList.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions));
                        jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(tempList.Select(c => new { SoPO = c.Manufa_Instruction_No, NgayXuat = c.Factory_Ship_Date }).ToList());
                    }
                }
                result = remainPT_PulleyLineA <= sum_totalInstruction_N1;
               
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return new Tuple<bool, double, int, string>(result, remainPT_PulleyLineA, sum_totalInstruction_N1, jsonStr);
        }

        //mode 2
        private Tuple<bool, double, int, string> CheckSPH_Mode2_CNC(int hour, List<OrderModel> lstOrders)
        {
            int sum_totalInstruction_N = 0;
            //int sum_totalInstruction_N1 = 0;

            bool result = false;
            string jsonStr = "";
            //2020-03-09 Start edit: 

            //lấy sum các giá trị của line pulley CNC
            double remainPT_PulleyLineA = OrderBLL.Instance.GetSumPulley_CNC(hour);

            //lấy các ngày lễ của năm hiện tại
            List<DateTime> lstHolidays = OrderBLL.Instance.GetHolidaysInYear();

            //2020-03-09 End edit
            
            var tempList = lstOrders.Where(d => lstHolidays.All(d2 => !d2.ToString("yyyyMMdd").Equals(d.Factory_Ship_Date.Value.ToString("yyyyMMdd"))) && d.Factory_Ship_Date.Value.ToString("yyyyMMdd") == DateTime.Now.ToString("yyyyMMdd")).ToList();
            if (tempList != null && tempList.Count > 0)
            {
                sum_totalInstruction_N = tempList.Sum(c => Convert.ToInt32(c.Number_of_Available_Instructions));
                jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(tempList.Select(c => new { SoPO = c.Manufa_Instruction_No, NgayXuat = c.Factory_Ship_Date }).ToList());
            }
            result = 0 < sum_totalInstruction_N;
            return new Tuple<bool, double, int, string>(result, remainPT_PulleyLineA, sum_totalInstruction_N, jsonStr);
        }
        private void SharingOrdersN1_CNC()
        {
            StringBuilder mailBody;
            //nếu máy CNC9, CNC10 có đơn N1 thì chia đều cho nhau
            
            int dayToAddN1 = HolidayHelper.Instance.DayCountN1;//nếu là thứ 7 thì lấy đơn N1 là ngày thứ 2

            //2019-12-16 Start add: tách lấy các đơn N+1 ra
            var ordersCNCHob9_N1 = ListOrdersCNCHob9.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();
            var ordersCNCHob10_N1 = ListOrdersCNCHob10.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.AddDays(dayToAddN1).Date).ToList();

            //CNC9 có N+1, CNC10 ko có
            if ((ordersCNCHob9_N1 != null && ordersCNCHob9_N1.Count > 0) && (ordersCNCHob10_N1 != null && ordersCNCHob10_N1.Count == 0))
            {
                //đếm số lượng của orders rồi chia 2
                var batchSize = Math.Ceiling(1.0 * ordersCNCHob9_N1.Count / 2);
                //chia đều đơn cho CNC9 và CNC10
                var ordersToSortCNC9 = ordersCNCHob9_N1.Take((int)batchSize).ToList();
                if(ordersToSortCNC9 != null && ordersToSortCNC9.Count > 0)
                {
                    var ordersForCNC9 = SortOrders_CNC(ordersToSortCNC9, Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                    //chèn các orders lần lượt trước các orders thường của các máy
                    ordersForCNC9.AddRange(ListOrdersCNCHob9);                   
                    //refresh CNC9, CNC10, remove duplicate
                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ordersForCNC9.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());

                    var ordersToSortCNC10 = ordersCNCHob9_N1.Except(ordersForCNC9).ToList();
                    if(ordersToSortCNC10 != null && ordersToSortCNC10.Count > 0)
                    {
                        var ordersForCNC10 = SortOrders_CNC(ordersToSortCNC10, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                        ordersForCNC10.AddRange(ListOrdersCNCHob10);
                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ordersForCNC10.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());
                    }

                }
              
            }//CNC10 có N+1, CNC9 ko có            
            else if ((ordersCNCHob9_N1 != null && ordersCNCHob9_N1.Count == 0) && (ordersCNCHob10_N1 != null && ordersCNCHob10_N1.Count > 0))
            {
                //đếm số lượng của orders rồi chia 2
                var batchSize = Math.Ceiling(1.0 * ordersCNCHob10_N1.Count / 2);
                //chia đều đơn cho CNC9 và CNC10
                var ordersToSortCNC10 = ordersCNCHob10_N1.Take((int)batchSize).ToList();
                if(ordersToSortCNC10!=null && ordersToSortCNC10.Count > 0)
                {
                    var ordersForCNC10 = SortOrders_CNC(ordersToSortCNC10, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                    //chèn các orders lần lượt trước các orders thường của các máy                   
                    ordersForCNC10.AddRange(ListOrdersCNCHob10);
                    //refresh CNC9, CNC10                   
                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ordersForCNC10.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());
                    
                    var ordersToSortCNC9 = ordersCNCHob10_N1.Except(ordersForCNC10).ToList();
                    if(ordersToSortCNC9 != null && ordersToSortCNC9.Count > 0)
                    {
                        var ordersForCNC9 = SortOrders_CNC(ordersToSortCNC9, Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                        ordersForCNC9.AddRange(ListOrdersCNCHob9);
                        ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ordersForCNC9.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());
                    }
                }                            
            }//CNC9 và CNC10 đều có N+1
            else if ((ordersCNCHob9_N1 != null && ordersCNCHob9_N1.Count > 0) && (ordersCNCHob10_N1 != null && ordersCNCHob10_N1.Count > 0))
            {
                //gộp đơn N+1 rồi chia đôi
                var ordersN1 = ordersCNCHob9_N1.Union(ordersCNCHob10_N1).ToList();
                //đếm số lượng của orders rồi chia 2
                var batchSize = Math.Ceiling(1.0 * ordersN1.Count / 2);
                //chia đều đơn cho CNC9 và CNC10
                var ordersToSortCNC9 = ordersN1.Take((int)batchSize).ToList();
                if(ordersToSortCNC9 != null && ordersToSortCNC9.Count > 0)
                {
                    var ordersForCNC9 = SortOrders_CNC(ordersToSortCNC9, Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                    //chèn các orders lần lượt trước các orders thường của các máy
                    ordersForCNC9.AddRange(ListOrdersCNCHob9);
                    //refresh CNC9, CNC10
                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ordersForCNC9.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());

                    var ordersToSortCNC10 = ordersN1.Except(ordersForCNC9).ToList();
                    if(ordersToSortCNC10 != null && ordersToSortCNC10.Count > 0)
                    {
                        var ordersForCNC10 = SortOrders_CNC(ordersToSortCNC10, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                        ordersForCNC10.AddRange(ListOrdersCNCHob10);
                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ordersForCNC10.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());
                    }                  
                }                             
            }
        }

        private void SharingOrdersN_CNC()
        {
            StringBuilder mailBody;
            //nếu máy CNC9, CNC10 có đơn N1 thì chia đều cho nhau

            //2019-12-16 Start add: tách lấy các đơn N+1 ra
            var ordersCNCHob9_N = ListOrdersCNCHob9.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();
            var ordersCNCHob10_N = ListOrdersCNCHob10.Where(c => c.Factory_Ship_Date.Value.Date == DateTime.Now.Date).ToList();

            //CNC9 có N, CNC10 ko có
            if ((ordersCNCHob9_N != null && ordersCNCHob9_N.Count > 0) && (ordersCNCHob10_N != null && ordersCNCHob10_N.Count == 0))
            {
                //đếm số lượng của orders rồi chia 2
                var batchSize = Math.Ceiling(1.0 * ordersCNCHob9_N.Count / 2);
                //chia đều đơn cho CNC9 và CNC10
                var ordersToSortCNC9 = ordersCNCHob9_N.Take((int)batchSize).ToList();
                if(ordersToSortCNC9 != null && ordersToSortCNC9.Count > 0)
                {
                    var ordersForCNC9 = SortOrders_CNC(ordersToSortCNC9, Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                    //chèn các orders lần lượt trước các orders thường của các máy
                    ordersForCNC9.AddRange(ListOrdersCNCHob9);
                    //refresh CNC9, CNC10
                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ordersForCNC9.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());

                    var ordersToSortCNC10 = ordersCNCHob9_N.Except(ordersForCNC9).ToList();
                    if(ordersToSortCNC10 != null && ordersToSortCNC10.Count >0)
                    {
                        var ordersForCNC10 = SortOrders_CNC(ordersToSortCNC10, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                        ordersForCNC10.AddRange(ListOrdersCNCHob10);
                        ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ordersForCNC10.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());
                    }                   
                }
                
            }//CNC10 có N+1, CNC9 ko có            
            else if ((ordersCNCHob9_N != null && ordersCNCHob9_N.Count == 0) && (ordersCNCHob10_N != null && ordersCNCHob10_N.Count > 0))
            {
                //đếm số lượng của orders rồi chia 2
                var batchSize = Math.Ceiling(1.0 * ordersCNCHob10_N.Count / 2);
                //chia đều đơn cho CNC9 và CNC10
                var ordersToSortCNC10 = ordersCNCHob10_N.Take((int)batchSize).ToList();
                if(ordersToSortCNC10 != null && ordersToSortCNC10.Count > 0)
                {
                    var ordersForCNC10 = SortOrders_CNC(ordersToSortCNC10, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                    ordersForCNC10.AddRange(ListOrdersCNCHob10);
                    //chèn các orders lần lượt trước các orders thường của các máy                    
                    //refresh CNC9, CNC10
                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ordersForCNC10.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());

                    var ordersToSortCNC9 = ordersCNCHob10_N.Except(ordersForCNC10).ToList();
                    if(ordersToSortCNC9 != null && ordersToSortCNC9.Count > 0)
                    {
                        var ordersForCNC9 = SortOrders_CNC(ordersToSortCNC9, Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                        ordersForCNC9.AddRange(ListOrdersCNCHob9);
                        ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ordersForCNC9.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());
                    }                  
                }
                
            }//CNC9 và CNC10 đều có N
            else if ((ordersCNCHob9_N != null && ordersCNCHob9_N.Count > 0) && (ordersCNCHob10_N != null && ordersCNCHob10_N.Count > 0))
            {
                //gộp đơn N rồi chia đôi
                var ordersN = ordersCNCHob9_N.Union(ordersCNCHob10_N).ToList();
                //đếm số lượng của orders rồi chia 2
                var batchSize = Math.Ceiling(1.0 * ordersN.Count / 2);
                //chia đều đơn cho CNC9 và CNC10
                var ordersToSortCNC9 = ordersN.Take((int)batchSize).ToList();
                if(ordersToSortCNC9 != null && ordersToSortCNC9.Count >0)
                {
                    var ordersForCNC9 = SortOrders_CNC(ordersToSortCNC9, Hob9CNC_DMat, Hob9CNC_TeethShape, Hob9CNC_TeethQty, Hob9CNC_Diameter);
                    //chèn các orders lần lượt trước các orders thường của các máy
                    ordersForCNC9.AddRange(ListOrdersCNCHob9);
                    //refresh CNC9, CNC10
                    ListOrdersCNCHob9 = new ObservableCollection<OrderModel>(ordersForCNC9.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());

                    var ordersToSortCNC10 = ordersN.Except(ordersForCNC9).ToList();
                    var ordersForCNC10 = SortOrders_CNC(ordersToSortCNC10, Hob10CNC_DMat, Hob10CNC_TeethShape, Hob10CNC_TeethQty, Hob10CNC_Diameter);
                    ordersForCNC10.AddRange(ListOrdersCNCHob10);
                    ListOrdersCNCHob10 = new ObservableCollection<OrderModel>(ordersForCNC10.GroupBy(c => c.Manufa_Instruction_No).Select(x => x.First()).ToList());
                }
            }
        }

        private void SharingOrdersCNC(string mode)
        {
            var t = Convert.ToInt32(DateTime.Now.ToString("HH"));
            if(t>=6 && t<24)//lúc này lấy orders N+1
            {
                SharingOrdersN1_CNC();
            }
            else if(t>=0 && t<6)
            {
                SharingOrdersN_CNC();//lấy orders N
            }
        }
        //2020-03-12 End add


        private void ApplySSDCheck_CNC(List<OrderModel> lstOrders)
        {
            var hour = Convert.ToInt32(DateTime.Now.ToString("HH"));
            //from 6:00 AM today to 2:00 AM tomorrow
            int[] hours_mode1 = { 6,7,8, 9,10,11, 12,13,14, 15,16,17, 18,19,20, 21,22,23, 0 };
            int[] hours_mode2 = { 2, 3, 4, 5 };
            int[] hourSendMail_mode1 = { 6, 9, 12, 15, 18, 21, 0 };
            int status = 0;
            try
            {
                //status = OrderBLL.Instance.Get_bwIsSetupN1();
                //check SSD mode1
                if (hours_mode1.Contains(hour))
                {
                    var resTup = CheckSPH_Mode1_CNC(hour, lstOrders);
                    //check SSD => true => cập nhật isBWCheckN1 true
                    if (resTup.Item1)
                    {
                        status = 1;
                        SSDModeCheck_CNC = "SSDMode";
                        CNC_RemainPT = resTup.Item2;
                        CNC_Total_N1_Orders = resTup.Item3;
                    }
                    else
                    {
                        status = 0;
                        SSDModeCheck_CNC = "Min Setup Mode";
                        CNC_RemainPT = resTup.Item2;
                        CNC_Total_N1_Orders = resTup.Item3;
                    }
                    if (isSendMailSSDModeCheck_CNC == false)
                    {
                        if(hourSendMail_mode1.Contains(hour))
                        {
                            StringBuilder builder = new StringBuilder();
                            builder.AppendLine("Mode check: " + SSDModeCheck_CNC + "</br>");
                            builder.AppendLine("RemainPT x SPH: " + resTup.Item2 + "</br>");
                            builder.AppendLine("N+1 Qty: " + resTup.Item3 + "</br>");
                            builder.AppendLine("PO List:</br>");
                            builder.AppendLine("------------------------------</br>");
                            builder.AppendLine(resTup.Item4);
                            var res = 1;//SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "thienphuoc@spclt.com.vn", "Thông báo SSDMode check CNC(1)", builder.ToString(), "");
                            isSendMailSSDModeCheck_CNC = res != 0 ? true : false;
                        }
                       
                    }
                }
                else if (hours_mode2.Contains(hour)) //check SSD mode2
                {
                    var resTup = CheckSPH_Mode2_CNC(hour, lstOrders);
                    //check SSD => true => cập nhật isCNCCheckN1 true
                    if (resTup.Item1)
                    {
                        status = 1;
                        SSDModeCheck_CNC = "SSDMode";
                        CNC_RemainPT = resTup.Item2;
                        CNC_Total_N1_Orders = resTup.Item3;
                    }
                    else
                    {
                        status = 0;
                        SSDModeCheck_CNC = "Min Setup Mode";
                        CNC_RemainPT = resTup.Item2;
                        CNC_Total_N1_Orders = resTup.Item3;
                    }
                    if (isSendMailSSDModeCheck_CNC == false)
                    {
                        StringBuilder builder = new StringBuilder();
                        builder.AppendLine("Mode check: " + SSDModeCheck_CNC + "</br>");
                        builder.AppendLine("RemainPT x SPH: " + resTup.Item2 + "</br>");
                        builder.AppendLine("N+1 Qty: " + resTup.Item3 + "</br>");
                        builder.AppendLine("PO List:</br>");
                        builder.AppendLine("------------------------------</br>");
                        builder.AppendLine(resTup.Item4);
                        var res = 1;//SPCMailHelper.Instance.PrepareSPCMail("honghanh@spclt.com.vn", "thienphuoc@spclt.com.vn", "Thông báo SSDMode check CNC(2)", builder.ToString(), "");
                        isSendMailSSDModeCheck_CNC = res != 0 ? true : false;
                    }
                }
                //else if(not_ssd_hours.Contains(hour))
                //{
                //    //hour not in hours_mode1 and hours_mode2
                //    //kiểm tra checkSPH => true tiếp tục ưu tiên N+1, cập nhật isBWCheckN1 true
                //    if(CheckSPH(hour,lstOrders))
                //    {
                //        status = 1;
                //        //if (status == 0)
                //        //{
                //        //    BalancedOrdersBW("SSD Check");
                //        //}
                //        //else
                //        //{
                //        //    BalancedOrdersBW_NormalMode();
                //        //}
                //    }
                //    else
                //    {
                //        //nếu check SPH false thì cân bằng order bình thường
                //        //BalancedOrdersBW_NormalMode();
                //        status = 0;
                //    }                              
                //}

                OrderBLL.Instance.Update_cncIsSetup(status);                
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        //2020-02-21 End add

        //2020-03-02 Start add: function format data in BW and CNC machine
        private string FormatHTMLOrders(List<OrderModel> source)
        {
            StringBuilder builder = new StringBuilder();
            List<object> lstOrders = new List<object>();
            foreach (var item in source)
            {
                lstOrders.Add(new { SoPO = item.Manufa_Instruction_No, TenHang = item.Item_Name, NgayNhan = item.Received_Date, NgayXuat = item.Factory_Ship_Date, SLDonHang = item.Number_of_Orders, SLHang = item.Number_of_Available_Instructions, Line = item.Line, MaVL = item.Material_text1, LoaiDao = item.TeethShape, SoRang = item.TeethQuantity, DuongKinh = item.Diameter_d });
            }

            return Newtonsoft.Json.JsonConvert.SerializeObject(lstOrders);
        }
        //2020-03-02 End add

        
        #endregion <<FUNCTIONS>>

        #region <<EVENTS>>
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }
        #endregion
    }

    enum SSDType
    {
        N1 = 1,
        N = 2

    }
}
