using BFDataCrawler.BLL;
using BFDataCrawler.Helper;
using BFDataCrawler.Helpers;
using BFDataCrawler.Model;
using mshtml;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Syroot.Windows.IO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;
using System.Windows.Threading;

namespace BFDataCrawler.ViewModel
{
    class OrderViewModel : INotifyPropertyChanged
    {

        #region CONSTRUCTOR
        public OrderViewModel()
        {
            Orders = new ObservableCollection<OrderModel>();
            //btnGetData
            IsLoading = false;
            IsGettingData = true;
            //btnPrintBWNoHob;
            BWNoHob_IsGettingData = true;
            BWNoHob_IsPrinting = false;
            //btnPrintBWHob
            BWHob_IsGettingData = true;
            BWHob_IsPrinting = false;

            //2019-10-10 Start add
            //clear all chromedrivers still exsist in TaskManager
            KillChromeDriver();
            //2019-10-10 End add


            //2019-10-14 Start add: auto synchronize deleted item from ListOrderBWHob
            timerBWHob = new DispatcherTimer()
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            timerBWHob.Start();
            timerBWHob.Tick += new EventHandler((s, e) => {
                //doc json va lay ra customer duoc delete tu process khac
               
                    //orderBWHob
                    OrderModel orderBWHob = JsonHelper.ReadJson2<OrderModel>(@"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWHob.json");
                    if (orderBWHob != null && ListOrdersBWHob != null)
                        ListOrdersBWHob.Remove(ListOrdersBWHob.FirstOrDefault(c => c.Manufa_Instruction_No.Equals(orderBWHob.Manufa_Instruction_No)));                   
             
            });

            timerCNCHob = new DispatcherTimer()
            {
                Interval = TimeSpan.FromSeconds(1)
            };

            timerCNCHob.Start();
            timerCNCHob.Tick += new EventHandler((s, e) => {
                //doc json va lay ra customer duoc delete tu process khac

                //orderCNCHob
                OrderModel orderCNCHob = JsonHelper.ReadJson2<OrderModel>(@"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob.json");
                if (orderCNCHob != null && ListOrdersCNCHob != null)
                    ListOrdersCNCHob.Remove(ListOrdersCNCHob.FirstOrDefault(c => c.Manufa_Instruction_No.Equals(orderCNCHob.Manufa_Instruction_No)));
                
            });
            //2019-10-14
        }
        #endregion

        #region PROPERTIES

        //2019-10-15 Start add: timer for synchronize deleted item from ListOrdersBWHob, ListOrdersCNCHob
        DispatcherTimer timerBWHob;
        DispatcherTimer timerCNCHob;
        //2019-10-15 End add

        bool isFirstNavigationBWNoHob = true;
        bool isFirstNavigationBWHob = true;
        //2019-10-07 Start add
        bool isFirstNavigationCNCNoHob = true;
        bool isFirstNavigationCNCHob = true;
        //2019-10-07 End add
        bool isFirstLoadBWNoHob = true;
        int pageNumberBWNoHob = 0;

        bool isFirstLoadBWHob = true;
        int pageNumberBWHob = 0;

        //2019-10-07 Start add
        bool isFirstLoadCNCNoHob = true;
        int pageNumberCNCNoHob = 0;

        bool isFirstLoadCNCHob = true;
        int pageNumberCNCHob = 0;
        //2019-10-07 End

        //2019-10-11 Start add
        bool _BWHob_IsOutOfSession = false;
        bool _BWNoHob_IsOutOfSession = false;
        bool _CNCHob_IsOutOfSession = false;
        bool _CNCNoHob_IsOutOfSession = false;
        //2019-10-11 End add

        //2019-10-15 Start add
        private bool timerIsRunning;
        public bool TimerIsRunning
        {
            get { return timerIsRunning; }
            set
            {
                timerIsRunning = value; OnPropertyChanged(nameof(TimerIsRunning));
            }
        }
        //2019-10-15 End add
        private bool _isLoading;
        public bool IsLoading
        {
            get { return _isLoading; }
            set
            {
                _isLoading = value;
                OnPropertyChanged("IsLoading");
            }
        }
        /// <summary>
        /// SelectedOrder is bind by selected row clicked on button in DataGrid via No property
        /// Get OrderInfo and bind to SelectedOrder then show it on OrderInfoUC for editing.
        /// 2019-07-30
        /// </summary>
        private OrderModel _selectedOrder;
        public OrderModel SelectedOrder
        {
            get { return _selectedOrder; }
            set
            {
                _selectedOrder = value;
                OnPropertyChanged("SelectedOrder");
            }
        }

        private ObservableCollection<OrderModel> _orders;
        public ObservableCollection<OrderModel> Orders
        {
            get { return _orders; }
            set { _orders = value; OnPropertyChanged("Orders"); }
        }


        private ObservableCollection<OrderModel> _orderBWHob;
        public ObservableCollection<OrderModel> ListOrdersBWHob
        {
            get { return _orderBWHob; }
            set
            {
                _orderBWHob = value;
                //2019-10-11 Start lock
                //if (_orderBWHob != null && _orderBWHob.Count < 6)
                //{
                //    //2019-10-09 Start add: send mail to SPC if orders are less than 6
                //    //mail body
                //    string mailContent = string.Format(@"
                //    Dear PC group, <br>
                //    The remaining available orders for <b>BW(Hob)</b> are <b>{0}</b>. Please make a preparation for new orders!<br>
                //    Thanks! <br><br>

                //    <b>Note:</b> This is announcement email, please do not reply! <br><br>
                //", _orderBWHob.Count);
                //    if (SPCMailHelper.Instance.PrepareSPCMail(SPCMailSender, "thienphuoc@spclt.com.vn", "Preparing new Manufa's orders for Pulley", mailContent) == 1)
                //        MailStatus = "A mail request for new orders has been sent!";
                //}
                //2019-10-09 End add
                //2019-10-11 End lock
                OnPropertyChanged(nameof(ListOrdersBWHob));
            }
        }

        private ObservableCollection<OrderModel> _orderBWNoHob;
        public ObservableCollection<OrderModel> ListOrdersBWNoHob
        {
            get { return _orderBWNoHob; }
            set {
                _orderBWNoHob = value;
                //2019-10-11 Start lock
                //if (_orderBWNoHob != null && _orderBWNoHob.Count < 6)
                //{
                //    //2019-10-09 Start add: send mail to SPC if orders are less than 6
                //    //mail body
                //    string mailContent = string.Format(@"
                //    Dear PC group, <br>
                //    The remaining available orders for <b>BW(NoHob)</b> are <b>{0}</b>. Please make a preparation for new orders!<br>
                //    Thanks! <br><br>

                //    <b>Note:</b> This is announcement email, please do not reply! <br><br>
                //", _orderBWNoHob.Count);
                //    if (SPCMailHelper.Instance.PrepareSPCMail(SPCMailSender, "thienphuoc@spclt.com.vn", "Preparing new Manufa's orders for Pulley", mailContent) == 1)
                //        MailStatus = "A mail request for new orders has been sent!";
                //}
                //2019-10-09 End add
                //2019-10-11 End lock
                OnPropertyChanged(nameof(ListOrdersBWNoHob)); }
        }

        private OrderModel _selectedOrderBWNoHob;
        public OrderModel SelectedOrderBWNoHob
        {
            get { return _selectedOrderBWNoHob; }
            set { _selectedOrderBWNoHob = value; OnPropertyChanged(nameof(SelectedOrderBWNoHob)); }
        }

        private OrderModel _selectedOrderBWHob;
        public OrderModel SelectedOrderBWHob
        {
            get { return _selectedOrderBWHob; }
            set { _selectedOrderBWHob = value; OnPropertyChanged(nameof(SelectedOrderBWHob)); }
        }


        //2019-10-05 Start add: new property
        private ObservableCollection<OrderModel> _orderCNCHob;
        public ObservableCollection<OrderModel> ListOrdersCNCHob
        {
            get { return _orderCNCHob; }
            set {
                _orderCNCHob = value;
                //2019-10-11 Start lock
                //if (_orderCNCHob != null && _orderCNCHob.Count < 6)
                //{
                //    //2019-10-09 Start add: send mail to SPC if orders are less than 6
                //    //mail body
                //    string mailContent = string.Format(@"
                //    Dear PC group, <br>
                //    The remaining available orders for <b>CNC(Hob)</b> are <b>{0}</b>. Please make a preparation for new orders!<br>
                //    Thanks! <br><br>

                //    <b>Note:</b> This is announcement email, please do not reply! <br><br>
                //", _orderCNCHob.Count);
                //    if (SPCMailHelper.Instance.PrepareSPCMail(SPCMailSender, "thienphuoc@spclt.com.vn", "Preparing new Manufa's orders for Pulley", mailContent) == 1)
                //        MailStatus = "A mail request for new orders has been sent!";
                //}
                //2019-10-09 End add
                //2019-10-11 End lock
                OnPropertyChanged(nameof(ListOrdersCNCHob)); }
        }

        private ObservableCollection<OrderModel> _orderCNCNoHob;
        public ObservableCollection<OrderModel> ListOrdersCNCNoHob
        {
            get { return _orderCNCNoHob; }
            set {
                _orderCNCNoHob = value;
                //2019-10-11 Start lock
                //if (_orderCNCNoHob != null && _orderCNCNoHob.Count < 6)
                //{
                //    //2019-10-09 Start add: send mail to SPC if orders are less than 6
                //    //mail body
                //    string mailContent = string.Format(@"
                //    Dear PC group, <br>
                //    The remaining available orders for <b>CNC(NoHob)</b> are <b>{0}</b>. Please make a preparation for new orders!<br>
                //    Thanks! <br><br>

                //    <b>Note:</b> This is announcement email, please do not reply! <br><br>
                //", _orderCNCNoHob.Count);
                //    if (SPCMailHelper.Instance.PrepareSPCMail(SPCMailSender, "thienphuoc@spclt.com.vn", "Preparing new Manufa's orders for Pulley", mailContent) == 1)
                //        MailStatus = "A mail request for new orders has been sent!";
                //}
                //2019-10-09 End add
                //2019-10-11 End lock
                OnPropertyChanged(nameof(ListOrdersCNCNoHob)); }
        }

        private OrderModel _selectedOrderCNCNoHob;
        public OrderModel SelectedOrderCNCNoHob
        {
            get { return _selectedOrderCNCNoHob; }
            set { _selectedOrderCNCNoHob = value; OnPropertyChanged(nameof(SelectedOrderCNCNoHob)); }
        }

        private OrderModel _selectedOrderCNCHob;
        public OrderModel SelectedOrderCNCHob
        {
            get { return _selectedOrderCNCHob; }
            set { _selectedOrderCNCHob = value; OnPropertyChanged(nameof(SelectedOrderCNCHob)); }
        }

        //2019-10-05 End add

        //2019-10-07 Start add: IsPrinting for BW and CNC print button
        private bool _BWHob_IsPrinting;
        public bool BWHob_IsPrinting
        {
            get { return _BWHob_IsPrinting; }
            set { _BWHob_IsPrinting = value; OnPropertyChanged(nameof(BWHob_IsPrinting)); }
        }

        private bool _BWNoHob_IsPrinting;
        public bool BWNoHob_IsPrinting
        {
            get { return _BWNoHob_IsPrinting; }
            set { _BWNoHob_IsPrinting = value; OnPropertyChanged(nameof(BWNoHob_IsPrinting)); }
        }

        private bool _CNCHob_IsPrinting;
        public bool CNCHob_IsPrinting
        {
            get { return _CNCHob_IsPrinting; }
            set { _CNCHob_IsPrinting = value; OnPropertyChanged(nameof(CNCHob_IsPrinting)); }
        }

        private bool _CNCNoHob_IsPrinting;
        public bool CNCNoHob_IsPrinting
        {
            get { return _CNCNoHob_IsPrinting; }
            set { _CNCNoHob_IsPrinting = value; OnPropertyChanged(nameof(CNCNoHob_IsPrinting)); }
        }

        //IsGettingData
        private bool _IsGettingData;
        public bool IsGettingData
        {
            get { return _IsGettingData; }
            set { _IsGettingData = value; OnPropertyChanged(nameof(IsGettingData)); }
        }
        //BWHob_IsGettingData
        private bool _BWHob_IsGettingData;
        public bool BWHob_IsGettingData
        {
            get { return _BWHob_IsGettingData; }
            set { _BWHob_IsGettingData = value; OnPropertyChanged(nameof(BWHob_IsGettingData)); }
        }
        //BWNoHob_IsGettingData
        private bool _BWNoHob_IsGettingData;
        public bool BWNoHob_IsGettingData
        {
            get { return _BWNoHob_IsGettingData; }
            set { _BWNoHob_IsGettingData = value; OnPropertyChanged(nameof(BWNoHob_IsGettingData)); }
        }
        //CNCHob_IsGettingData
        private bool _CNCHob_IsGettingData;
        public bool CNCHob_IsGettingData
        {
            get { return _CNCHob_IsGettingData; }
            set { _CNCHob_IsGettingData = value; OnPropertyChanged(nameof(CNCHob_IsGettingData)); }
        }
        //CNCNoHob_IsGettingData
        private bool _CNCNoHob_IsGettingData;
        public bool CNCNoHob_IsGettingData
        {
            get { return _CNCNoHob_IsGettingData; }
            set { _CNCNoHob_IsGettingData = value; OnPropertyChanged(nameof(CNCNoHob_IsGettingData)); }
        }

        //2019-10-08 Start add

        public string HobBW_TeethShape
        {
            get { return Properties.Settings.Default._hobBWTeethShape; }
            set
            {
                Properties.Settings.Default._hobBWTeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(HobBW_TeethShape));
            }
        }

        public int HobBW_TeethQty
        {
            get { return Properties.Settings.Default._hobBWTeethQty; }
            set
            {
                Properties.Settings.Default._hobBWTeethQty = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(HobBW_TeethQty));
            }
        }

        public string HobCNC_TeethShape
        {
            get { return Properties.Settings.Default._hobCNCTeethShape; }
            set
            {
                Properties.Settings.Default._hobCNCTeethShape = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(HobCNC_TeethShape));
            }
        }

        public int HobCNC_TeethQty
        {
            get { return Properties.Settings.Default._hobCNCTeethQty; }
            set
            {
                Properties.Settings.Default._hobCNCTeethQty = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(HobCNC_TeethQty));
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
        //2019-10-08 End add

        //2019-10-07 End add

        //2019-10-09 Start add
        private string _mailStatus;
        public string MailStatus
        {
            get { return _mailStatus; }
            set
            {
                _mailStatus = value;
                OnPropertyChanged(nameof(MailStatus));
            }
        }

        
        public string SPCMailSender
        {
            get {
                return Properties.Settings.Default._spcMailSender;
            }
            set {
                Properties.Settings.Default._spcMailSender = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(SPCMailSender)); }
        }

        
        public string SPCMailReceiver
        {
            get { return Properties.Settings.Default._spcMailReceiver; }
            set {
                Properties.Settings.Default._spcMailReceiver = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(SPCMailReceiver)); }
        }

        
        public string SPCMailCode
        {
            get { return Properties.Settings.Default._spcMailCode; }
            set {
                Properties.Settings.Default._spcMailCode = value;
                Properties.Settings.Default.Save();
                OnPropertyChanged(nameof(SPCMailCode)); }
        }
        //2019-10-09 End add

        //2019-10-11 Start add
        private string _BW_HobStatus;
        public string BW_HobStatus
        {
            get { return _BW_HobStatus; }
            set { _BW_HobStatus = value; OnPropertyChanged(nameof(BW_HobStatus)); }
        }

        private string _BW_NoHobStatus;
        public string BW_NoHobStatus
        {
            get { return _BW_NoHobStatus; }
            set { _BW_NoHobStatus = value; OnPropertyChanged(nameof(BW_NoHobStatus)); }
        }

        private string _CNC_HobStatus;
        public string CNC_HobStatus
        {
            get { return _CNC_HobStatus; }
            set { _CNC_HobStatus = value; OnPropertyChanged(nameof(CNC_HobStatus)); }
        }

        private string _CNC_NoHobStatus;
        public string CNC_NoHobStatus
        {
            get { return _CNC_NoHobStatus; }
            set { _CNC_NoHobStatus = value; OnPropertyChanged(nameof(CNC_NoHobStatus)); }
        }
        //2019-10-11 End add
        #endregion

        #region Commands
        private ICommand _commandGetOrders;
        public ICommand CommandGetOrders
        {
            get
            {
                if (_commandGetOrders == null)
                    _commandGetOrders = new RelayCommand<object>(CanGetOrders, GetOrders);

                return _commandGetOrders;
            }
            private set { _commandGetOrders = value; }
        }

        //2019-10-04 Start add
        private ICommand _commandBWPrintHob;
        public ICommand CommandBWPrintHob
        {
            get
            {
                if (_commandBWPrintHob == null)
                    _commandBWPrintHob = new RelayCommand<object>(CanPrintBWHob, PrintBWHob);
                return _commandBWPrintHob;
            }
            private set { _commandBWPrintHob = value; }
        }
        //2019-10-04 End add


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


        //2019-10-07 Start add
        private ICommand _commandCNCPrintHob;
        public ICommand CommandCNCPrintHob
        {
            get
            {
                if (_commandCNCPrintHob == null)
                    _commandCNCPrintHob = new RelayCommand<object>(CanPrintCNCHob, PrintCNCHob);
                return _commandCNCPrintHob;
            }
            private set { _commandCNCPrintHob = value; }
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
        //2019-10-07 End add
        //2019-10-14 Start add: Test delete and generate sample data
        private ICommand _commandGenerateData;
        public ICommand CommandGenerateData
        {
            get {
                if (_commandGenerateData == null)
                    _commandGenerateData = new RelayCommand<object>(GenerateData);
                return _commandGenerateData;
            }
            private set { _commandGenerateData = value; }
        }
        private void GenerateData(object obj)
        {            
            ObservableCollection<OrderModel> lstBWHobForTest = new ObservableCollection<OrderModel>();
          
            for (int i = 0; i < 10; i++)
            {
               
                lstBWHobForTest.Add(new OrderModel()
                {
                    
                    Sales_Order_No = "BW",
                    Pier_Instruction_No = "",
                    Manufa_Instruction_No = "A"+i+"-TEST",
                    Global_Code = "",
               
                    Item_Name = "",
                    MC = "",
                    Received_Date = DateTime.Now,
                    Factory_Ship_Date = DateTime.Now,
                    Number_of_Orders = "",
                    Number_of_Available_Instructions = "",
               
                    Line = "",
               
                    Material_code1 = "",
                    Material_text1 = "",
                    Amount_used1 = "",
                    Unit1 = "",
                 
                    Inner_Code = "",
                  
                    Hobbing = true,
                    TeethQuantity = 10,
                    TeethShape = "",
                    LineAB = "BW"
                });
            }

            ListOrdersBWHob = new ObservableCollection<OrderModel>(lstBWHobForTest);
        }

        private ICommand _commandDeleteBWHob;
        public ICommand CommandDeleteBWHob
        {
            get
            {
                if (_commandDeleteBWHob == null)
                    _commandDeleteBWHob = new RelayCommand<object>(DeleteBWHob);
                return _commandDeleteBWHob;
            }
            private set { _commandDeleteBWHob = value; }
        }

        private void DeleteBWHob(object obj)
        {
            if (ListOrdersBWHob != null)
            {
                SelectedOrderBWHob = ListOrdersBWHob.FirstOrDefault();
                if(SelectedOrderBWHob != null)
                {
                    JsonHelper.SaveJson2(SelectedOrderBWHob, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWHob.json");
                    ListOrdersBWHob.Remove(SelectedOrderBWHob);
                }                               
            }
        }

        //2019-10-14 End add
        //2019-10-04 Start add
        private async void PrintBWHob(object parameter)
        {
            if (parameter != null)
            {
                var values = parameter as object[];
                WebBrowser wb = values[0] as WebBrowser;
                mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

                DataGrid gv = values[1] as DataGrid;
                await Task.Run(() =>
                {

                    IsGettingData = false;
                    //2019-10-15 Start add: stop timer while PO is printing
                    timerBWHob.Stop();
                    //2019-10-15 End add
                    SelectedOrderBWHob = ListOrdersBWHob[0];
                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                    btnClear.click();
                });

            }
        }

        private bool CanPrintBWHob(object obj)
        {
            if (ListOrdersBWHob != null && ListOrdersBWHob.Count > 0 && IsLoading == false)
                return true;
            else return false;
        }
        //2019-10-04 End add

        private async void PrintBWNoHob(object obj)
        {
            if (obj != null)
            {
                WebBrowser wb = obj as WebBrowser;
                mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

                await Task.Run(() =>
                {
                    //var pierInstructionNo = document.getElementById("ContentPlaceHolder1_txtPierAufnr");
                    //pierInstructionNo.innerText = SelectedOrder.Manufa_Instruction_No;
                    //pierInstructionNo.innerText = "101000135085";

                    //var manufaInstructionNo = document.getElementById("ContentPlaceHolder1_txtAufnr");
                    //manufaInstructionNo.innerHTML = SelectedOrderBWNoHob.Manufa_Instruction_No;
                    //var btnSearch = document.getElementById("ContentPlaceHolder1_cmbSelect");
                    IsGettingData = false;
                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                    btnClear.click();
                    //btnSearch.click();


                });

            }
        }

        private bool CanPrintBWNoHob(object obj)
        {
            if (SelectedOrderBWNoHob != null && IsLoading == false)
                return true;
            else return false;

        }

        //2019-10-07 Start add

        private async void PrintCNCHob(object parameter)
        {
            if (parameter != null)
            {
                var values = parameter as object[];
                WebBrowser wb = values[0] as WebBrowser;
                mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

                DataGrid gv = values[1] as DataGrid;
                await Task.Run(() =>
                {
                    IsGettingData = false;
                    //2019-10-15 Start add: stop timer while PO is printing
                    timerCNCHob.Stop();
                    //2019-10-15 End add
                    SelectedOrderCNCHob = ListOrdersCNCHob[0];
                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                    btnClear.click();
                });

            }
        }

        private bool CanPrintCNCHob(object obj)
        {
            if (ListOrdersCNCHob != null && ListOrdersCNCHob.Count > 0 && IsLoading == false)
                return true;
            else return false;
        }

        private bool CanPrintCNCNoHob(object obj)
        {
            if (SelectedOrderCNCNoHob != null && IsLoading == false)
                return true;
            else return false;
        }

        private async void PrintCNCNoHob(object obj)
        {
            if (obj != null)
            {
                WebBrowser wb = obj as WebBrowser;
                mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;

                await Task.Run(() =>
                {
                    IsGettingData = false;
                    var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                    btnClear.click();
                });

            }
        }

        private bool CanGetOrders(object obj)
        {
            //if the followed printing statuses TRUE then the button is Disable
            if (BWHob_IsPrinting || BWNoHob_IsPrinting || CNCHob_IsPrinting || CNCNoHob_IsPrinting)
                return false;
            else return true;

        }
        //2019-10-07 End add
        private async void GetOrders(object parameters)
        {
            if (parameters != null)
            {
                IsLoading = true;
                //2019-10-07 Start lock
                //bt.IsEnabled = IsLoading ? false : true;
                //2019-10-07 End lock

                //2019-10-07 Start add
                IsGettingData = false;
                //2019-10-07 End add
              

                var values = parameters as object[];
                Button bt = values[0] as Button;
                WebBrowser wbBWNoHob = values[1] as WebBrowser;
                //2019-10-04 Start add
                WebBrowser wbBWHob = values[2] as WebBrowser;
                //2019-10-04 End add

                //2019-10-07 Start add
                WebBrowser wbCNCNoHob = values[3] as WebBrowser;
                WebBrowser wbCNCHob = values[4] as WebBrowser;
                //2019-10-07 End add

                //get user login info in App.config
                string user = ConfigurationManager.AppSettings["user"];
                string pass = ConfigurationManager.AppSettings["pass"];

                //hide command prompt and browser when chromeDriver start
                var options = new ChromeOptions();
                options.AddArgument("--window-position=-32000,-32000");                
                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
               
                
               
               
                ChromeDriver chromeDriver = new ChromeDriver(service, options);
                
                await Task.Run(() =>
                {
                    chromeDriver.Url = "http://10.4.24.111:8080/manufa/Login.aspx";
                    chromeDriver.Navigate();

                    if (isFirstNavigationBWNoHob)
                    {
                        //navigate BWNoHob page
                        wbBWNoHob.Dispatcher.InvokeAsync(() =>
                        {
                            wbBWNoHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                            wbBWNoHob.LoadCompleted += new LoadCompletedEventHandler(wbBWNoHobLoadCompleted);
                        });
                    }

                    if (isFirstNavigationBWHob)
                    {
                        //navigate BWHob page
                        wbBWHob.Dispatcher.InvokeAsync(() =>
                        {
                            wbBWHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                            wbBWHob.LoadCompleted += new LoadCompletedEventHandler(wbBWHobLoadCompleted);
                        });
                    }

                    //2019-10-07 Start add
                    if (isFirstNavigationCNCNoHob)
                    {
                        //navigate CNCNoHob page
                        wbCNCNoHob.Dispatcher.InvokeAsync(() =>
                        {
                            wbCNCNoHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                            wbCNCNoHob.LoadCompleted += new LoadCompletedEventHandler(wbCNCNoHob_LoadCompleted);
                        });
                    }

                    if (isFirstNavigationCNCHob)
                    {
                        //navigate CNCHob page
                        wbCNCHob.Dispatcher.InvokeAsync(() =>
                        {
                            wbCNCHob.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                            wbCNCHob.LoadCompleted += new LoadCompletedEventHandler(wbCNCHob_LoadCompleted);
                        });
                    }
                    //2019-10-07 End add
                });


               
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
                    catch (Exception)
                    {
                        MessageBox.Show("OrdersList excel file is being used, please close it and and try again!", "Message", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }


                    //login info
                    var userInput = chromeDriver.FindElementById("txtUserId");
                    var passInput = chromeDriver.FindElementById("txtPassword");
                    var btnLogin = chromeDriver.FindElementById("cmdLogin");
                    //login input
                    userInput.SendKeys(user);
                    passInput.SendKeys(pass);
                    btnLogin.Click();
                    //navigate manufacturing
                    var manufacturing_Instructions_Print = chromeDriver.FindElementByXPath("//*[@id=\"MenuDetail_0\"]/tbody/tr[1]/td");
                    manufacturing_Instructions_Print.Click();

                    Task.Delay(1000).Wait();
                    //navigate orders list
                    var forms = chromeDriver.FindElementById("ContentPlaceHolder1_drpSelectForm");
                    SelectElement selectForm = new SelectElement(forms);
                    selectForm.SelectByValue("Orders List");

                    Task.Delay(1000).Wait();
                    //input search filter
                    var statuses = chromeDriver.FindElementById("ContentPlaceHolder1_drpStatus");
                    SelectElement selectStatus = new SelectElement(statuses);
                    selectStatus.SelectByValue("Authorization already (Not Print)");

                    //input search filter
                    var line = chromeDriver.FindElementById("ContentPlaceHolder1_txtLine");
                    line.SendKeys("LE");
                    //onclick search
                    var btnSearch = chromeDriver.FindElementById("ContentPlaceHolder1_cmbSelect");
                    btnSearch.Click();

                    //download file excel
                    var btnExcelOutput = chromeDriver.FindElementById("ContentPlaceHolder1_cmbCsvOutput");
                    btnExcelOutput.Click();

                    Task.Delay(5000).Wait();
                    //fetch rows to get orders
                    //Orders = FetchOrders(chromeDriver);

                    //dispose chromeDriver 
                    chromeDriver.Close();
                    chromeDriver.Quit();
                    
                    //read file Excel
                    string excelFile = downloadFolderPath + "\\OrdersList.xlsx";
                    List<OrderModel> lstManufaInfo = GetOrdersFromExcel(excelFile).ToList();
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
                            //2019-10-05 Start lock
                            //IEnumerable<OrderModel> result = (from r1 in lstManufaInfo
                            //                                  join
                            //         r2 in lstPulleyMaster on r1.Inner_Code equals r2.Inner_Code
                            //                                  select new OrderModel()
                            //                                  {
                            //                                      //No = r1.No,
                            //                                      Sales_Order_No = r1.Sales_Order_No,
                            //                                      Pier_Instruction_No = r1.Pier_Instruction_No,
                            //                                      Manufa_Instruction_No = r1.Manufa_Instruction_No,
                            //                                      Global_Code = r1.Global_Code,
                            //                                      //Customers = r1.Customers,
                            //                                      Item_Name = r1.Item_Name,
                            //                                      MC = r1.MC,
                            //                                      Received_Date = r1.Received_Date,
                            //                                      Factory_Ship_Date = r1.Factory_Ship_Date,
                            //                                      Number_of_Orders = r1.Number_of_Orders,
                            //                                      Number_of_Available_Instructions = r1.Number_of_Available_Instructions,
                            //                                      //Number_of_Repairs = r1.Number_of_Repairs,
                            //                                      //Number_of_Instructions = r1.Number_of_Instructions,
                            //                                      Line = r1.Line,
                            //                                      //PayWard = r1.PayWard,
                            //                                      //Major = r1.Major,
                            //                                      //Special_Orders = r1.Special_Orders,
                            //                                      //Method = r1.Method,
                            //                                      //Destination = r1.Destination,
                            //                                      //Instructions_Print_date = r1.Instructions_Print_date,
                            //                                      //Latest_progress = r1.Latest_progress,
                            //                                      //Tack_Label_Output_Date = r1.Tack_Label_Output_Date,
                            //                                      //Completion_Instruction_Date = r1.Completion_Instruction_Date,
                            //                                      //Re_print_Count = r1.Re_print_Count,
                            //                                      //Latest_issue_time = r1.Latest_issue_time,
                            //                                      Material_code1 = r1.Material_code1,
                            //                                      Material_text1 = r1.Material_text1,
                            //                                      Amount_used1 = r1.Amount_used1,
                            //                                      Unit1 = r1.Unit1,
                            //                                      //Material_code2 = r1.Material_code2,
                            //                                      //Material_text2 = r1.Material_text2,
                            //                                      //Amount_used2 = r1.Amount_used2,
                            //                                      //Unit2 = r1.Unit2,
                            //                                      //Material_code3 = r1.Material_code3,
                            //                                      //Material_text3 = r1.Material_text3,
                            //                                      //Amount_used3 = r1.Amount_used3,
                            //                                      //Unit3 = r1.Unit3,
                            //                                      //Material_code4 = r1.Material_code4,
                            //                                      //Material_text4 = r1.Material_text4,
                            //                                      //Amount_used4 = r1.Amount_used4,
                            //                                      //Unit4 = r1.Unit4,
                            //                                      //Material_code5 = r1.Material_code5,
                            //                                      //Material_text5 = r1.Material_text5,
                            //                                      //Amount_used5 = r1.Amount_used5,
                            //                                      //Unit5 = r1.Unit5,
                            //                                      //Material_code6 = r1.Material_code6,
                            //                                      //Material_text6 = r1.Material_text6,
                            //                                      //Amount_used6 = r1.Amount_used6,
                            //                                      //Unit6 = r1.Unit6,
                            //                                      //Material_code7 = r1.Material_code7,
                            //                                      //Material_text7 = r1.Material_text7,
                            //                                      //Amount_used7 = r1.Amount_used7,
                            //                                      //Unit7 = r1.Unit7,
                            //                                      //Material_code8 = r1.Material_code8,
                            //                                      //Material_text8 = r1.Material_text8,
                            //                                      //Amount_used8 = r1.Amount_used8,
                            //                                      //Unit8 = r1.Unit8,
                            //                                      Inner_Code = r1.Inner_Code,
                            //                                      //Classify_Code = r1.Classify_Code,
                            //                                      Hobbing = r2.Hobbing,
                            //                                      TeethQuantity = r2.TeethQuantity,
                            //                                      TeethShape = r2.TeethShape
                            //                                  }).ToList();
                            //2019-10-05 End lock

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
                                                                  Material_text1 = r1.Material_text1,
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
                                                                  LineAB = r3.LineAB
                                                              }).ToList();

                            Orders = new ObservableCollection<OrderModel>(result);
                            //get Hobbing, teethQuantity, teethShape
                            //BW orders with Hob
                            if (HobBW_TeethQty > 0 && !HobBW_TeethShape.Equals(""))
                            {
                                ListOrdersBWHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing && c.LineAB.Equals("BW")).ToList().OrderByDescending(c => HobBW_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }
                            else
                            {
                                ListOrdersBWHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing && c.LineAB.Equals("BW")).ToList().OrderByDescending(c => HobBW_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }


                            //BW orders with No Hob
                            if (NoHobBW_TeethQty > 0 && !NoHobBW_TeethShape.Equals(""))
                            {
                                ListOrdersBWNoHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing == false && c.LineAB.Equals("BW")).ToList().OrderByDescending(c => NoHobBW_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }
                            else
                            {
                                ListOrdersBWNoHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing == false && c.LineAB.Equals("BW")).ToList().OrderByDescending(c => NoHobBW_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }


                            //CNC Hob
                            if (HobCNC_TeethQty > 0 && !HobCNC_TeethShape.Equals(""))
                            {
                                ListOrdersCNCHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => HobCNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }
                            else
                            {
                                ListOrdersCNCHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => HobCNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }



                            //CNC NoHob
                            if (NoHobCNC_TeethQty > 0 && !NoHobCNC_TeethShape.Equals(""))
                            {
                                ListOrdersCNCNoHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing == false && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => NoHobCNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }
                            else
                            {
                                ListOrdersCNCNoHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing == false && c.LineAB.Equals("CNC")).ToList().OrderByDescending(c => NoHobCNC_TeethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity));
                            }

                        }

                    }



                    //if (Orders.Count > 0)
                    //{
                    //    //add orders to DB
                    //    int result = OrderBLL.Instance.AddNewOrders(Orders.ToList());
                    //    if (result > 0)
                    //        MessageBox.Show(string.Format(@"{0} order(s) successfully added!", result), "Message", MessageBoxButton.OK, MessageBoxImage.Information);
                    //}

                    IsLoading = false;
                    //2019-10-09 Start add
                    IsGettingData = true;
                    //2019-10-09 End add

                    //2019-10-09 Start lock
                    //bt.Dispatcher.Invoke(() =>
                    //{
                    //    bt.IsEnabled = IsLoading ? false : true;
                    //});
                    //2019-10-09 End lock
                });


            }

        }



        private ICommand _commandEditOrderInfo;
        public ICommand CommandEditOrderInfo
        {
            get
            {
                if (_commandEditOrderInfo == null)
                    _commandEditOrderInfo = new RelayCommand<object>(EditOrderInfo);
                return _commandEditOrderInfo;
            }
            private set { _commandEditOrderInfo = value; OnPropertyChanged("CommandEditOrderInfo"); }
        }

        private void EditOrderInfo(object obj)
        {
            throw new NotImplementedException();
        }

        private ICommand _cmdCrawlManufaData;
        public ICommand CmdCrawlManufaData
        {
            get
            {
                if (_cmdCrawlManufaData == null)
                    _cmdCrawlManufaData = new RelayCommand<object>(CrawlManufaData);

                return _cmdCrawlManufaData;
            }
            private set { _cmdCrawlManufaData = value; OnPropertyChanged(nameof(CmdCrawlManufaData)); }
        }

        private ICommand _cmdSortOrderInfo;
        public ICommand CmdSortOrderInfo
        {
            get
            {
                if (_cmdSortOrderInfo == null)
                    _cmdSortOrderInfo = new RelayCommand<object>(SortOrderInfo);

                return _cmdSortOrderInfo;
            }
            private set
            {
                _cmdSortOrderInfo = value;
            }
        }

        private void SortOrderInfo(object parameters)
        {
            if (parameters != null)
            {
                var values = (object[])parameters;

                int teethQuantity = string.IsNullOrEmpty(values[0].ToString()) ? 0 : Convert.ToInt32(values[0]);
                string teethShape = values[1].ToString();
                //priority teethShape -> teethQuantity -> bare (not included yet)
                if (!string.IsNullOrEmpty(teethShape) && teethQuantity >= 0)
                {
                    var result = Orders.OrderByDescending(c => c.Hobbing).ThenByDescending(c => teethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity).SkipWhile(c => c.Hobbing == false);
                    Orders = new ObservableCollection<OrderModel>(result);
                    //BW orders with Hob
                    //2019-10-03 Start delete
                    //ListOrdersBWHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing).
                    //                                                                      Select(c => new OrderModel()
                    //                                                                      {
                    //                                                                          Sales_Order_No = c.Sales_Order_No,
                    //                                                                          Pier_Instruction_No = c.Pier_Instruction_No,
                    //                                                                          Manufa_Instruction_No = c.Manufa_Instruction_No,
                    //                                                                          Global_Code = c.Global_Code,
                    //                                                                          Item_Name = c.Item_Name,
                    //                                                                          MC = c.MC,
                    //                                                                          Received_Date = c.Received_Date,
                    //                                                                          Factory_Ship_Date = c.Factory_Ship_Date,
                    //                                                                          Number_of_Orders = c.Number_of_Orders,
                    //                                                                          Number_of_Available_Instructions = c.Number_of_Available_Instructions,
                    //                                                                          Line = c.Line,
                    //                                                                          Material_text1 = c.Material_text1,
                    //                                                                          Inner_Code = c.Inner_Code,
                    //                                                                          Hobbing = c.Hobbing,
                    //                                                                          TeethQuantity = c.TeethQuantity,
                    //                                                                          TeethShape = c.TeethShape

                    //                                                                      }).ToList());
                    //2019-10-03 End delete
                    ListOrdersBWHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing));
                    //BW orders with No Hob
                    ListOrdersBWNoHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing == false));
                }
                else
                {
                    //Hobbing:Yes/No -> DESC: teethShape -> ASC: teethQty
                    var result = Orders.OrderByDescending(c => c.Hobbing).ThenByDescending(c => teethShape.Equals(c.TeethShape)).ThenBy(c => c.TeethQuantity).ThenBy(c => c.TeethShape).ThenBy(c => c.TeethQuantity).SkipWhile(c => c.Hobbing == false);
                    Orders = new ObservableCollection<OrderModel>(result);
                    ListOrdersBWHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing));
                    //BW orders with No Hob
                    ListOrdersBWNoHob = new ObservableCollection<OrderModel>(Orders.Where(c => c.Hobbing == false));
                }


                //Orders = new ObservableCollection<OrderModel>(sortedOrders);
            }
        }

        private void CrawlManufaData(object obj)
        {
            if (obj != null)
            {
                page = 0;
                WebBrowser wb = obj as WebBrowser;

                if (wb!=null)
                    wb.LoadCompleted -= wbBW_LoadCompleted;
                //WebBrowser wb = new WebBrowser();
                wb.Navigate("http://10.4.24.111:8080/manufa/Login.aspx");
                //2019-10-30 Lock
                //wb.LoadCompleted += new LoadCompletedEventHandler(wbBWNoHobLoadCompleted);
                //2019-10-30 End
                
                //2019-10-30 Start add
                wb.LoadCompleted += new LoadCompletedEventHandler(wbBW_LoadCompleted);
                //2019-10-30 End add
            }
        }

        int page = 0;
        private void wbBW_LoadCompleted(object sender, NavigationEventArgs e)
        {
            WebBrowser wb = sender as WebBrowser;
            mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;
            if(page == 0)
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
                page = 1;
            }
            else if(page == 1)
            {
                if(wb.IsLoaded)
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
                page = 2;
            }
            else if(page == 2)
            {
                if(wb.IsLoaded)
                {
                    var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                    var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                    script.text = @" 
                                   window.onload = function(){
                                        var select = document.getElementById('ContentPlaceHolder1_drpStatus');
                                        select.selectedIndex = 2; //order not printed yet
                                        document.getElementById('ContentPlaceHolder1_txtLine').value = 'SH';

                                        var displayItem = document.getElementById('ContentPlaceHolder1_drpDispCount');                                        
                                        displayItem.selectedIndex = 1;//choose 500 items                                        

                                        var btnSearch = document.getElementById('ContentPlaceHolder1_cmbSelect');
                                        btnSearch.click();
                                        
                                        //setTimeout(function(){
                                        //    var btnExcel = document.getElementById('ContentPlaceHolder1_cmbCsvOutput');
                                        //    btnExcel.click();
                                        //},5000);

                                   }
                              ";
                    head.appendChild((mshtml.IHTMLDOMNode)script);
                    HideScriptErrors(wb, true);

                    
                }
                page = 3;

            }
            else if(page == 3)
            {
                if(wb.IsLoaded)
                {
                    //2019-10-31 Lock
                    //var head = document.getElementsByTagName("body").Cast<mshtml.HTMLBody>().First();

                    //var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                    //script.text = @" 
                    //               //window.onload = function(){
                    //               //    getTable();                                      
                    //               //}

                    //                function getTable()
                    //                {
                    //                    //var table = document.getElementById('ContentPlaceHolder1_sprPrintList_viewport').outerHTML;
                    //                    //table.querySelector('tbody').outerHTML;
                    //                    var table = document.getElementById('ContentPlaceHolder1_sprPrintList_viewport');
                    //                    return table.querySelector('tbody').outerHTML;
                    //                }
                    //          ";
                    //head.appendChild((mshtml.IHTMLDOMNode)script);
                    //HideScriptErrors(wb, true);
                    //var table = wb.InvokeScript("getTable");
                    //string htmlTable = string.Format(@"<table id='mytable'>{0}</table>", table.ToString());
                    //2019-10-31 End lock





                    //lấy html cả trang web
                    //var htmlPage = document.documentElement.outerHTML;
                    var htmltb = document.getElementById("ContentPlaceHolder1_sprPrintList_viewport").outerHTML;
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();                    
                    doc.LoadHtml(htmltb);

                    List<OrderModel> lst = new List<OrderModel>();

                    //var rows = doc.DocumentNode.SelectNodes("//*[@id='ContentPlaceHolder1_sprPrintList_viewport']/tbody/tr").FirstOrDefault().SelectNodes("//*[@id='ContentPlaceHolder1_sprPrintList_viewport']/tbody/tr[1]/td");
                    
                    
                    var rows = doc.DocumentNode.SelectNodes("//*[@id='ContentPlaceHolder1_sprPrintList_viewport']/tbody/tr");
                    DateTime validDatetime;
                    for (int i = 0; i < rows.Count; i++)
                    {
                        //var cols = rows[i].SelectNodes("//*[@id='ContentPlaceHolder1_sprPrintList_viewport']/tbody/tr["+i+"]/td");
                        var cols = rows[i].SelectNodes("//*[@id='ContentPlaceHolder1_sprPrintList_viewport']/tbody/tr["+(i+1)+"]/td");
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
                                Inner_Code = cols[58].InnerText.ToString() ?? "",
                                Classify_Code = cols[59].InnerText.ToString() ?? "",
                            });
                        }

                    }

                    Orders = new ObservableCollection<OrderModel>(lst);

                    var lnkLogout = document.getElementById("lnkLogOut");
                    lnkLogout.click();
                    
                }
                page = 4;
            }
            else if(page == 4)
            {
                //if (wb != null)
                //{                   
                //    wb.Navigate((Uri)null);
                //}

                page = 0;//quay về
            }
            else
            {

            }
        }

        //properties for getting cookie from webserver
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetGetCookieEx(string pchURL, string pchCookieName, StringBuilder pchCookieData, ref uint pcchCookieData, int dwFlags, IntPtr lpReserved);


        //2019-10-04 Start add
        //KEY-BWHob
        private void wbBWHobLoadCompleted(object sender, NavigationEventArgs e)
        {
            isFirstNavigationBWHob = false;
            WebBrowser wb = sender as WebBrowser;
            mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;
            if (isFirstLoadBWHob)
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
                isFirstLoadBWHob = false;
            }
            else
            {
                //access 2nd time
                //using HTML Agility pack to get html content
                //string url2nd = "http://10.4.24.111:8080/manufa/MainMenu.aspx";
                //var web = new HtmlAgilityPack.HtmlWeb();
                //var doc = web.Load(url2nd);
                //var instructionPrintNode = doc.DocumentNode.SelectNodes("//*[@id=\"MenuDetail_0\"]/tbody/tr[1]/td").FirstOrDefault();
                //var html = new HtmlAgilityPack.HtmlDocument();
                //html.LoadHtml(document.documentElement.innerHTML);
                //var table = html.GetElementbyId("MenuDetail_0");
                //var rows = table.Element("tbody").Elements("tr");   
                //   var rows = document.getElementsByTagName("tr").Cast<mshtml.HTMLElementCollection>();
                if (pageNumberBWHob == 0)
                {
                    //nếu trang tải xong
                    if (wb.IsLoaded)
                    {
                        //check session time out
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadBWHob = true; //true: login page <> false: other pages     
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
                                        window.location = 'PrintViewForm.aspx?DB=1';
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);

                            pageNumberBWHob = 1;
                        }
                    }

                }
                else if (pageNumberBWHob == 1)
                {
                    //check session time out
                    if (CheckSessionExpired(wb))
                    {
                        //navigate to login page
                        isFirstLoadBWHob = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();
                    }
                    else
                    {
                        var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                        var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                        script.text = @" 
                                   window.onload = function(){
                                       
                                        //select Manufacturing Instructions Print option to navigate new page
                                        window.location = '/manufa/PrintViewInstruction.aspx';                                                  
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        HideScriptErrors(wb, true);
                        pageNumberBWHob = 2;
                    }

                }
                else if (pageNumberBWHob == 2)
                {
                   
                    //check session time out
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if BWHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _BWHob_IsOutOfSession = true;
                        //2019-10-11 End add

                        //navigate to login page
                        isFirstLoadBWHob = true; //true: login page <> false: other pages     
                        document.getElementById("btnOK").click();
                        
                    }
                    else
                    {
                        //2019-10-11 Start add
                        if (_BWHob_IsOutOfSession)
                        {
                            _BWHob_IsOutOfSession = false;//reset to default value
                            SelectedOrderBWHob = ListOrdersBWHob[0];
                            //onclick btnClear to navigate to page 3
                            pageNumberBWHob = 3;
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
                            pageNumberBWHob = 3;

                        }

                    }
                }
                else if (pageNumberBWHob == 3)
                {
                    //Disable btnGetData, enable BWHobLoading while the printer is running                    
                    BWHob_IsGettingData = false;
                    BWHob_IsPrinting = true;
                    if (wb.IsLoaded)
                    {
                        
                        if (CheckSessionExpired(wb))
                        {
                            //2019-10-11 Start add: if BWHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                            _BWHob_IsOutOfSession = true;
                            //2019-10-11 End add

                            //navigate to login page
                            isFirstLoadBWHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();
                            
                        }
                        else
                        {                           
                                //điền PO (manufa NO) và "Search" đơn để in
                                document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWHob.Manufa_Instruction_No);
                                document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                                pageNumberBWHob = 4;                                                  
                        }
                    }


                }
                else if (pageNumberBWHob == 4)
                {
                    if (CheckSessionExpired(wb))
                    {
                        //2019-10-11 Start add: if BWHob run out of Session then re-login and auto simulate click on btnClear in Manufa's website and re-print the order
                        _BWHob_IsOutOfSession = true;
                        //2019-10-11 End add

                        //navigate to login page
                        isFirstLoadBWHob = true; //true: login page <> false: other pages     
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
                        pageNumberBWHob = 5; //chuyển trang sau khi in, lấy status
                    }

                }
                else if (pageNumberBWHob == 5)
                {
                    if (wb.IsLoaded)
                    {
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadBWHob = true; //true: login page <> false: other pages     
                            document.getElementById("btnOK").click();
                                                   
                        }
                        else
                        {
                            //lấy status sau khi in
                            string status = document.getElementById("ContentPlaceHolder1_AlertArea").innerHTML;
                            //nếu in thành công
                            if (status.ToLower().Contains("completed"))
                            {
                                

                                //2019-10-08 Start add
                                HobBW_TeethShape = SelectedOrderBWHob.TeethShape;
                                HobBW_TeethQty = SelectedOrderBWHob.TeethQuantity;
                                //2019-10-08 End add

                                //xóa PO đã in thành công khỏi list và lưu item đã xóa vào json cho việc đồng bộ dữ  liệu khi sử dụng nhiều máy tính với nhau
                                var poToDelete = ListOrdersBWHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                {
                                    ListOrdersBWHob.Remove(poToDelete);                                    
                                    JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\BWHob.json");
                                    //2019-10-15 Start add: enable timer again
                                    timerBWHob.Start();
                                    //2019-10-15 End add
                                }
                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add

                                //2019-10-07 Start add
                                //Enable btnGetData, disable BWHobLoading while the printer is running
                                IsGettingData = true;
                                BWHob_IsGettingData = true;
                                BWHob_IsPrinting = false;
                                //2019-10-07 End add
                                pageNumberBWHob = 1;//quay lại từ đầu
                            }
                            else if (status.ToLower().Contains("Print Target"))
                            {
                                //xóa PO đã in thành công khỏi list
                                var poToDelete = ListOrdersBWHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                {
                                    ListOrdersBWHob.Remove(poToDelete);

                                }
                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add

                                //2019-10-11 Start add: print message: Order has been printed
                                BW_HobStatus = "The selected order has been printed from other machine!";
                                //2019-10-07 Start add
                                //Enable btnGetData, disable BWHobLoading while the printer is running
                                IsGettingData = true;
                                BWHob_IsGettingData = true;
                                BWHob_IsPrinting = false;
                                //2019-10-07 End add
                                pageNumberBWHob = 1;//quay lại từ đầu
                            }
                            else
                            {
                                //máy in hết giấy, lỗi data, etc..
                                //quay về trang in
                                //2019-10-07 Start add
                                //Enable btnGetData, disable BWHobLoading if getting error
                                IsGettingData = true;
                                BWHob_IsGettingData = true;
                                BWHob_IsPrinting = false;
                                //2019-10-07 End add

                                pageNumberBWHob = 1;
                            }

                        }

                    }

                }
                else if (pageNumberBWHob == 6)
                {

                }
                else { }

            }
        }
        //2019-10-04 End add
        //KEY-BWNoHob
        private void wbBWNoHobLoadCompleted(object sender, NavigationEventArgs e)
        {
            isFirstNavigationBWNoHob = false;
            WebBrowser wb = sender as WebBrowser;
            mshtml.HTMLDocument document = (mshtml.HTMLDocument)wb.Document;
            if (isFirstLoadBWNoHob)
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
                isFirstLoadBWNoHob = false;
            }
            else
            {
                //access 2nd time
                //using HTML Agility pack to get html content
                //string url2nd = "http://10.4.24.111:8080/manufa/MainMenu.aspx";
                //var web = new HtmlAgilityPack.HtmlWeb();
                //var doc = web.Load(url2nd);
                //var instructionPrintNode = doc.DocumentNode.SelectNodes("//*[@id=\"MenuDetail_0\"]/tbody/tr[1]/td").FirstOrDefault();
                //var html = new HtmlAgilityPack.HtmlDocument();
                //html.LoadHtml(document.documentElement.innerHTML);
                //var table = html.GetElementbyId("MenuDetail_0");
                //var rows = table.Element("tbody").Elements("tr");   
                //   var rows = document.getElementsByTagName("tr").Cast<mshtml.HTMLElementCollection>();
                if (pageNumberBWNoHob == 0)
                {
                    //nếu trang tải xong
                    if (wb.IsLoaded)
                    {
                        //check session time out
                        if (CheckSessionExpired(wb))
                        {
                            //navigate to login page
                            isFirstLoadBWNoHob = true; //true: login page <> false: other pages
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
                                        window.location = 'PrintViewForm.aspx?DB=1';
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);

                            pageNumberBWNoHob = 1;
                        }
                    }

                }
                else if (pageNumberBWNoHob == 1)
                {
                    //check session time out
                    if (CheckSessionExpired(wb))
                    {
                        //navigate to login page
                        isFirstLoadBWNoHob = true; //true: login page <> false: other pages
                        document.getElementById("btnOK").click();
                    }
                    else
                    {
                        var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                        var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                        script.text = @" 
                                   window.onload = function(){
                                       
                                        //select Manufacturing Instructions Print option to navigate new page
                                        window.location = '/manufa/PrintViewInstruction.aspx';                                                  
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        HideScriptErrors(wb, true);
                        pageNumberBWNoHob = 2;
                    }

                }
                else if (pageNumberBWNoHob == 2)
                {
                    //var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                    //var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                    //script.text = @" 
                    //               window.onload = function(){                                                                              
                    //                    //set selected status option for the order that not printed yet
                    //                    var statusSelect = document.getElementById('ContentPlaceHolder1_drpStatus');
                    //                    statusSelect.selectedIndex = 2;

                    //                    //search order for PULLEY line
                    //                    var txtLine = document.getElementById('ContentPlaceHolder1_txtLine');
                    //                    txtLine.innerText = 'LE'; 

                    //                    //perform button Search
                    //                    var btnSearch = document.getElementById('ContentPlaceHolder1_cmbSelect');
                    //                    btnSearch.click();

                    //               }
                    //          ";
                    //head.appendChild((mshtml.IHTMLDOMNode)script);
                    //check session time out
                    if (CheckSessionExpired(wb))
                    {
                        //navigate to login page
                        isFirstLoadBWNoHob = true; //true: login page <> false: other pages
                        document.getElementById("btnOK").click();                                             
                    }
                    else
                    {
                        //2019-10-11 Start add
                        if (_BWNoHob_IsOutOfSession)
                        {
                            _BWNoHob_IsOutOfSession = false;//reset to default value                           
                            //onclick btnClear to navigate to page 3
                            pageNumberBWNoHob = 3;
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
                            pageNumberBWNoHob = 3;
                        }
                        
                    }


                }
                else if (pageNumberBWNoHob == 3)
                {
                    //Disable btnGetData, enable BWNoHobLoading while the printer is running

                    BWNoHob_IsGettingData = false;
                    BWNoHob_IsPrinting = true;
                    if (wb.IsLoaded)
                    {
                        if (CheckSessionExpired(wb))
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
                            //điền PO (manufa NO) và "Search" đơn để in
                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderBWNoHob.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            pageNumberBWNoHob = 4;
                        }
                    }


                }
                else if (pageNumberBWNoHob == 4)
                {
                    if (CheckSessionExpired(wb))
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
                                            
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },5000);
                                            //in 
                                            
                                        }
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        pageNumberBWNoHob = 5; //chuyển trang sau khi in, lấy status
                    }

                }
                else if (pageNumberBWNoHob == 5)
                {
                    if (wb.IsLoaded)
                    {
                        if (CheckSessionExpired(wb))
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
                                //nếu in thành công, lưu TeethShape, TeethQty, Hop của PO vừa in thành công để so sánh và sắp xếp cho list
                                //2019-10-08 Start add
                                NoHobBW_TeethShape = SelectedOrderBWNoHob.TeethShape;
                                NoHobBW_TeethQty = SelectedOrderBWNoHob.TeethQuantity;
                                //2019-10-08 End add

                                //xóa PO đã in thành công khỏi list
                                var poToDelete = ListOrdersBWNoHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWNoHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                    ListOrdersBWNoHob.Remove(poToDelete);

                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add

                                //2019-10-07 Start add
                                //Enable btnGetData, disable BWNoHobLoading when the process finished
                                IsGettingData = true;
                                BWNoHob_IsGettingData = true;
                                BWNoHob_IsPrinting = false;
                                //2019-10-07 End add

                                pageNumberBWNoHob = 1;//quay lại từ đầu
                            }
                            else if (status.ToLower().Contains("Print Target"))
                            {
                                //xóa PO đã in thành công khỏi list
                                var poToDelete = ListOrdersBWNoHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderBWNoHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                    ListOrdersBWNoHob.Remove(poToDelete);

                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add

                                //2019-10-11 Start add: print message: Order has been printed
                                BW_NoHobStatus = "The selected order has been printed from other machine!";
                                //2019-10-11 End add

                                //2019-10-07 Start add
                                //Enable btnGetData, disable BWNoHobLoading when the process finished
                                IsGettingData = true;
                                BWNoHob_IsGettingData = true;
                                BWNoHob_IsPrinting = false;
                                //2019-10-07 End add

                                pageNumberBWNoHob = 1;//quay lại từ đầu
                            }
                            else
                            {
                                //máy in hết giấy, lỗi data, etc..
                                //quay về trang in

                                //2019-10-07 Start add
                                //Enable btnGetData, disable BWNoHobLoading when the process finished
                                IsGettingData = true;
                                BWNoHob_IsGettingData = true;
                                BWNoHob_IsPrinting = false;
                                //2019-10-07 End add
                                pageNumberBWNoHob = 1;
                            }

                        }

                    }

                }
                else if (pageNumberBWNoHob == 6)
                {

                }
                else { }

            }

        }

        //2019-10-07 Start add
        //KEY-CNCHob
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
                                        window.location = 'PrintViewForm.aspx?DB=1';
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);

                            pageNumberCNCHob = 1;
                        }
                    }

                }
                else if (pageNumberCNCHob == 1)
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
                        var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                        var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                        script.text = @" 
                                   window.onload = function(){
                                       
                                        //select Manufacturing Instructions Print option to navigate new page
                                        window.location = '/manufa/PrintViewInstruction.aspx';                                                  
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        HideScriptErrors(wb, true);
                        pageNumberCNCHob = 2;
                    }

                }
                else if (pageNumberCNCHob == 2)
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
                        //2019-10-11 Start add
                        if (_CNCHob_IsOutOfSession)
                        {
                            _CNCHob_IsOutOfSession = false;//reset to default value
                            SelectedOrderCNCHob = ListOrdersCNCHob[0];
                            //onclick btnClear to navigate to page 3
                            pageNumberCNCHob = 3;
                            var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                            btnClear.click();
                        }
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
                            pageNumberCNCHob = 3;
                        }
                        
                    }


                }
                else if (pageNumberCNCHob == 3)
                {
                    //Disable btnGetData, enable CNCHobLoading while the printer is running                    
                    CNCHob_IsGettingData = false;
                    CNCHob_IsPrinting = true;
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
                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCHob.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            pageNumberCNCHob = 4;
                        }
                    }


                }
                else if (pageNumberCNCHob == 4)
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
                                            
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },2000);
                                            //in 
                                            
                                        }
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        pageNumberCNCHob = 5; //chuyển trang sau khi in, lấy status
                    }

                }
                else if (pageNumberCNCHob == 5)
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
                                //nếu in thành công, lưu TeethShape, TeethQty, Hop của PO vừa in thành công để so sánh và sắp xếp cho list
                                //2019-10-08 Start add
                                HobCNC_TeethShape = SelectedOrderCNCHob.TeethShape;
                                HobCNC_TeethQty = SelectedOrderCNCHob.TeethQuantity;
                                //2019-10-08 End add
                                //xóa PO đã in thành công khỏi list
                                var poToDelete = ListOrdersCNCHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                {
                                    JsonHelper.SaveJson2(poToDelete, @"\\10.4.17.62\F4_App\Programs\BWDataCrawler\CNCHob.json");
                                    ListOrdersCNCHob.Remove(poToDelete);
                                    //2019-10-15 Start add: start timer again
                                    timerCNCHob.Start();
                                    //2019-10-15
                                }
                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();

                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCHobLoading when the process finished
                                IsGettingData = true;
                                CNCHob_IsGettingData = true;
                                CNCHob_IsPrinting = false;
                                //2019-10-07 End add

                                //2019-10-04 End add
                                pageNumberCNCHob = 1;//quay lại từ đầu
                            }
                            else if (status.ToLower().Contains("Print Target"))
                            {
                                //xóa PO đã in thành công khỏi list
                                var poToDelete = ListOrdersCNCHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                    ListOrdersCNCHob.Remove(poToDelete);

                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add

                                //2019-10-11 Start add: print message: Order has been printed
                                CNC_HobStatus = "The selected order has been printed from other machine!";
                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCHobLoading when the process finished
                                IsGettingData = true;
                                CNCHob_IsGettingData = true;
                                CNCHob_IsPrinting = false;
                                //2019-10-07 End add

                                //2019-10-04 End add
                                pageNumberCNCHob = 1;//quay lại từ đầu
                            }
                            else
                            {

                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCHobLoading when the process finished
                                IsGettingData = true;
                                CNCHob_IsGettingData = true;
                                CNCHob_IsPrinting = false;
                                //2019-10-07 End add
                                //máy in hết giấy, lỗi data, etc..
                                //quay về trang in
                                pageNumberCNCHob = 1;
                            }

                        }

                    }

                }
                else if (pageNumberCNCHob == 6)
                {

                }
                else { }

            }
        }
        //KEY-CNCNoHob
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
                                        window.location = 'PrintViewForm.aspx?DB=1';
                                   }
                              ";
                            head.appendChild((mshtml.IHTMLDOMNode)script);

                            pageNumberCNCNoHob = 1;
                        }
                    }

                }
                else if (pageNumberCNCNoHob == 1)
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
                        var head = document.getElementsByTagName("head").Cast<mshtml.HTMLHeadElement>().First();

                        var script = (mshtml.IHTMLScriptElement)document.createElement("script");
                        script.text = @" 
                                   window.onload = function(){
                                       
                                        //select Manufacturing Instructions Print option to navigate new page
                                        window.location = '/manufa/PrintViewInstruction.aspx';                                                  
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        HideScriptErrors(wb, true);
                        pageNumberCNCNoHob = 2;
                    }

                }
                else if (pageNumberCNCNoHob == 2)
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
                        //2019-10-11 Start add
                        if (_CNCNoHob_IsOutOfSession)
                        {
                            _CNCNoHob_IsOutOfSession = false;//reset to default value                           
                            //onclick btnClear to navigate to page 3
                            pageNumberCNCNoHob = 3;
                            var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                            btnClear.click();
                        }//2019-10-11 End add
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
                            pageNumberCNCNoHob = 3;
                        }
                       
                    }


                }
                else if (pageNumberCNCNoHob == 3)
                {
                    //Disable btnGetData, enable CNCNoHobLoading while the printer is running                    
                    CNCNoHob_IsGettingData = false;
                    CNCNoHob_IsPrinting = true;
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
                            document.getElementById("ContentPlaceHolder1_txtAufnr").setAttribute("value", SelectedOrderCNCNoHob.Manufa_Instruction_No);
                            document.getElementById("ContentPlaceHolder1_cmbSelect").click();
                            pageNumberCNCNoHob = 4;
                        }
                    }


                }
                else if (pageNumberCNCNoHob == 4)
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
                                            
                                            setTimeout(function(){
                                                btnPrint.click();
                                            },5000);
                                            //in 
                                            
                                        }
                                   }
                              ";
                        head.appendChild((mshtml.IHTMLDOMNode)script);
                        pageNumberCNCNoHob = 5; //chuyển trang sau khi in, lấy status
                    }

                }
                else if (pageNumberCNCNoHob == 5)
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
                                //nếu in thành công, lưu TeethShape, TeethQty, Hop của PO vừa in thành công để so sánh và sắp xếp cho list
                                //2019-10-08 Start add
                                NoHobCNC_TeethShape = SelectedOrderCNCNoHob.TeethShape;
                                NoHobCNC_TeethQty = SelectedOrderCNCNoHob.TeethQuantity;
                                //2019-10-08 End add

                                //xóa PO đã in thành công khỏi list
                                var poToDelete = ListOrdersCNCNoHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                    ListOrdersCNCNoHob.Remove(poToDelete);

                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();

                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCNoHobLoading when the process finished
                                IsGettingData = true;
                                CNCNoHob_IsGettingData = true;
                                CNCNoHob_IsPrinting = false;
                                //2019-10-07 End add

                                //2019-10-04 End add
                                pageNumberCNCNoHob = 1;//quay lại từ đầu
                            }
                            else if (status.ToLower().Contains("Print Target"))
                            {
                                //xóa PO đã in thành công khỏi list
                                var poToDelete = ListOrdersCNCNoHob.SingleOrDefault(c => c.Manufa_Instruction_No == SelectedOrderCNCNoHob.Manufa_Instruction_No);
                                if (poToDelete != null)
                                    ListOrdersCNCNoHob.Remove(poToDelete);

                                //2019-10-04 Start add: clear status sau khi in và chuyển hướng về bước 1
                                var btnClear = document.getElementById("ContentPlaceHolder1_cmbClear");
                                btnClear.click();
                                //2019-10-04 End add

                                //2019-10-11 Start add: print message: Order has been printed
                                CNC_NoHobStatus = "The selected order has been printed from other machine!";
                                //2019-10-11 End add

                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCNoHobLoading when the process finished
                                IsGettingData = true;
                                CNCNoHob_IsGettingData = true;
                                CNCNoHob_IsPrinting = false;
                                //2019-10-07 End add

                                //2019-10-04 End add
                                pageNumberCNCNoHob = 1;//quay lại từ đầu
                            }
                            else
                            {
                                //2019-10-07 Start add
                                //Enable btnGetData, disable CNCNoHobLoading when the process finished
                                IsGettingData = true;
                                CNCNoHob_IsGettingData = true;
                                CNCNoHob_IsPrinting = false;
                                //2019-10-07 End add

                                //máy in hết giấy, lỗi data, etc..
                                //quay về trang in
                                pageNumberCNCNoHob = 1;
                            }

                        }

                    }

                }
                else if (pageNumberCNCNoHob == 6)
                {

                }
                else { }

            }
        }
        //2019-19-07 End add

        #endregion


        #region FUNCTIONS
        public string GetHTMLPage(string url)
        {
            Uri uri = new Uri(url);
            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string html = reader.ReadToEnd();
            return html;
        }

        private static void DownloadCurrent()
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create("http://10.4.24.111:8080/manufa/DownloadPage_common.aspx?FileName=OrdersList");
            webRequest.Method = "GET";
            webRequest.Timeout = 3000;
            webRequest.BeginGetResponse(new AsyncCallback(PlayResponeAsync), webRequest);
        }

        private static void PlayResponeAsync(IAsyncResult asyncResult)
        {
            long total = 0;
            long received = 0;
            HttpWebRequest webRequest = (HttpWebRequest)asyncResult.AsyncState;

            try
            {
                using (HttpWebResponse webResponse = (HttpWebResponse)webRequest.EndGetResponse(asyncResult))
                {
                    byte[] buffer = new byte[1024];

                    FileStream fileStream = File.OpenWrite(@"D:\test\OrderList.xlsx");
                    using (Stream input = webResponse.GetResponseStream())
                    {
                        total = webResponse.ContentLength;

                        int size = input.Read(buffer, 0, buffer.Length);
                        while (size > 0)
                        {
                            fileStream.Write(buffer, 0, size);
                            received += size;

                            size = input.Read(buffer, 0, buffer.Length);
                        }
                    }

                    fileStream.Flush();
                    fileStream.Close();
                }
            }
            catch (Exception ex)
            {
            }
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

        private ObservableCollection<OrderModel> GetOrdersFromExcel(string excelFile)
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

        private ObservableCollection<OrderModel> FetchOrders(ChromeDriver chromeDriver)
        {
            ObservableCollection<OrderModel> lst = new ObservableCollection<OrderModel>();
            try
            {
                IWebElement tableElement = chromeDriver.FindElementById("ContentPlaceHolder1_sprPrintList_viewport");
                IList<IWebElement> tableRow = tableElement.FindElements(By.TagName("tr"));
                IList<IWebElement> rowTD;



                if (tableRow.Count > 0)
                {
                    DateTime dt = new DateTime(2000, 1, 1);//default date

                    foreach (var row in tableRow)
                    {
                        rowTD = row.FindElements(By.TagName("td"));
                        lst.Add(new OrderModel()
                        {
                            No = rowTD[0].Text,
                            Sales_Order_No = rowTD[1].Text,
                            Pier_Instruction_No = rowTD[2].Text,
                            Manufa_Instruction_No = rowTD[3].Text,
                            Global_Code = rowTD[4].Text,

                            Customers = rowTD[5].Text,
                            Item_Name = rowTD[6].Text,
                            MC = rowTD[7].Text,
                            Received_Date = DateTime.TryParseExact(rowTD[8].Text, "yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt) ? Convert.ToDateTime(rowTD[8].Text) : dt,
                            Factory_Ship_Date = DateTime.TryParseExact(rowTD[9].Text, "yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt) ? Convert.ToDateTime(rowTD[9].Text) : dt,
                            Number_of_Orders = rowTD[10].Text,
                            Number_of_Available_Instructions = rowTD[11].Text,
                            Number_of_Repairs = rowTD[12].Text,
                            Number_of_Instructions = rowTD[13].Text,
                            Line = rowTD[14].Text,
                            PayWard = rowTD[15].Text,
                            Major = rowTD[16].Text,
                            Special_Orders = rowTD[17].Text,
                            Method = rowTD[18].Text,
                            Destination = rowTD[19].Text,
                            Instructions_Print_date = DateTime.TryParseExact(rowTD[20].Text, "yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt) ? Convert.ToDateTime(rowTD[20].Text) : dt,
                            Latest_progress = rowTD[21].Text,
                            Tack_Label_Output_Date = rowTD[22].Text,
                            Completion_Instruction_Date = DateTime.TryParseExact(rowTD[23].Text, "yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt) ? Convert.ToDateTime(rowTD[23].Text) : dt,
                            Re_print_Count = rowTD[24].Text,
                            Latest_issue_time = rowTD[25].Text,
                            //History1 = rowTD[26].Text,
                            //History2 = rowTD[27].Text,
                            //History3 = rowTD[28].Text,
                            //History4 = rowTD[29].Text,
                            Material_code1 = rowTD[26].Text,
                            Material_text1 = rowTD[27].Text,
                            Amount_used1 = rowTD[28].Text,
                            Unit1 = rowTD[29].Text,
                            Material_code2 = rowTD[30].Text,
                            Material_text2 = rowTD[31].Text,
                            Amount_used2 = rowTD[32].Text,
                            Unit2 = rowTD[33].Text,
                            Material_code3 = rowTD[34].Text,
                            Material_text3 = rowTD[35].Text,
                            Amount_used3 = rowTD[36].Text,
                            Unit3 = rowTD[37].Text,
                            Material_code4 = rowTD[38].Text,
                            Material_text4 = rowTD[39].Text,
                            Amount_used4 = rowTD[40].Text,
                            Unit4 = rowTD[41].Text,
                            Material_code5 = rowTD[42].Text,
                            Material_text5 = rowTD[43].Text,
                            Amount_used5 = rowTD[44].Text,
                            Unit5 = rowTD[45].Text,
                            Material_code6 = rowTD[46].Text,
                            Material_text6 = rowTD[47].Text,
                            Amount_used6 = rowTD[48].Text,
                            Unit6 = rowTD[49].Text,
                            Material_code7 = rowTD[50].Text,
                            Material_text7 = rowTD[51].Text,
                            Amount_used7 = rowTD[52].Text,
                            Unit7 = rowTD[53].Text,
                            Material_code8 = rowTD[54].Text,
                            Material_text8 = rowTD[55].Text,
                            Amount_used8 = rowTD[56].Text,
                            Unit8 = rowTD[57].Text,
                            Inner_Code = rowTD[58].Text,
                            Classify_Code = rowTD[59].Text

                        });
                    }


                }
                chromeDriver.Close();
                chromeDriver.Quit();
                return lst;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private bool CheckSessionExpired(WebBrowser browser)
        {
            if (browser.Source.AbsoluteUri.Equals("http://10.4.24.111:8080/error/SessionTimeOut.aspx"))
                return true;
            else return false;
        }

        //2019-10-10 Start add: function to kill processes named "chromedriver"
        private void KillChromeDriver()
        {
            foreach (var process in Process.GetProcessesByName("chromedriver"))
            {
                process.Kill();
            }
        }
        //2019-10-10 End add
        #endregion

        #region EVENTS
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }
        #endregion

    }
}
