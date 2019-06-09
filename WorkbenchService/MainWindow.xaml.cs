using System;
using System.Windows;
using System.IO.Ports;
using System.Collections.Generic;
using System.Windows.Input;
using System.Text;
using System.Threading;
using System.Collections;
using System.Security.Permissions;
using System.Windows.Threading;
using System.Net;
using System.IO;
using System.Timers;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace WorkbenchService
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        SerialPort ComPort = new SerialPort();//声明一个串口
        private string[] ports;//可用串口数组
        IList<customer> comList = new List<customer>();//可用串口集合
        private bool WaitClose = false;//invoke里判断是否正在关闭串口是否正在关闭串口，执行Application.DoEvents，并阻止再次invoke ,解决关闭串口时，程序假死，具体参见http://news.ccidnet.com/art/32859/20100524/2067861_4.html 仅在单线程收发使用，但是在公共代码区有相关设置，所以未用#define隔离
        Queue recQueue = new Queue();//接收数据过程中，接收数据线程与数据处理线程直接传递的队列，先进先出
        private bool Listening = false;//用于检测是否没有执行完invoke相关操作，仅在单线程收发使用，但是在公共代码区有相关设置，所以未用#define隔离
        private bool ComPortIsOpen = false;//COM口开启状态字，在打开/关闭串口中使用，这里没有使用自带的ComPort.IsOpen，因为在串口突然丢失的时候，ComPort.IsOpen会自动false，逻辑混乱
        private Byte[] Terminals = new Byte[9];//保存发送指令
        string Url;//获取网址
        private System.Timers.Timer AskTerminalTask = new System.Timers.Timer();//发送指令到终端的定时任务
        private System.Timers.Timer FileSaveTask = new System.Timers.Timer();//保存文件的定时任务
        private int CollectCount = 0;//采集次数
        private int CollectIntervalTime = 0;//采集间隔时间
        private int PollingIntervalTime = 0;//轮询间隔时间
        private int FileSavePeriod = 0;//文件保存周期
        private string JSONSaveDocument = "C:\\JSONSave";

        //终端编号
        Byte TerminalNumber = 0;
        //空气湿度
        Byte AtmosphereHumidity_H, AtmosphereHumidity_L;
        UInt16 AtmosphereHumidity;
        //空气温度
        Byte AtmosphereTemperature_H, AtmosphereTemperature_L;
        UInt16 AtmosphereTemperature;
        //光照强度
        Byte Illumination_01, Illumination_02, Illumination_03, Illumination_04;
        UInt32 Illumination;
        //土壤湿度
        Byte SoilHumidity_H, SoilHumidity_L;
        UInt16 SoilHumidity01, SoilHumidity02, SoilHumidity03, SoilHumidity04, SoilHumidity05, SoilHumidity06, SoilHumidity07, SoilHumidity08, SoilHumidity09;
        //土壤温度
        Byte SoilTemperature_H, SoilTemperature_L;
        UInt16 SoilTemperature01, SoilTemperature02, SoilTemperature03, SoilTemperature04, SoilTemperature05, SoilTemperature06, SoilTemperature07, SoilTemperature08, SoilTemperature09;
        //土壤酸碱度
        Byte SoliPH_H, SOliPH_L;
        UInt16 SoliPH;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)//主窗口初始化    
        {
            Thread _TimeShow = new Thread(CurrentTimeShow);//显示当前时间
            _TimeShow.Start();

            SendDataTextBox.IsReadOnly = true;
            SystemMessageTextBox.IsReadOnly = true;

            //↓↓↓↓↓↓↓↓↓指令赋值↓↓↓↓↓↓↓↓↓
            Terminals[0] = 0x01;
            Terminals[1] = 0x0d;
            Terminals[2] = 0x0a;
            Terminals[3] = 0x02;
            Terminals[4] = 0x0d;
            Terminals[5] = 0x0a;
            Terminals[6] = 0x03;
            Terminals[7] = 0x0d;
            Terminals[8] = 0x0a;
            //↑↑↑↑↑↑↑↑↑指令赋值↑↑↑↑↑↑↑↑↑

            //↓↓↓↓↓↓↓↓↓文件操作↓↓↓↓↓↓↓↓↓
            if (!Directory.Exists(JSONSaveDocument))
            {
                Directory.CreateDirectory(JSONSaveDocument);
            }
            //↑↑↑↑↑↑↑↑↑文件操作↑↑↑↑↑↑↑↑↑

            //↓↓↓↓↓↓↓↓↓可用串口下拉控件↓↓↓↓↓↓↓↓↓
            ports = SerialPort.GetPortNames();//获取可用串口
            if (ports.Length > 0)//ports.Length > 0说明有串口可用
            {
                for (int i = 0; i < ports.Length; i++)
                {
                    comList.Add(new customer() { com = ports[i] });//下拉控件里添加可用串口
                }
                AvailableComCbobox.ItemsSource = comList;//资源路劲
                AvailableComCbobox.DisplayMemberPath = "com";//显示路径
                AvailableComCbobox.SelectedValuePath = "com";//值路径
                AvailableComCbobox.SelectedValue = ports[0];//默认选第1个串口
            }
            else//未检测到串口
            {
                SystemMessageTextBox.Text = "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "无可用串口";
                SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
            }
            //↑↑↑↑↑↑↑↑↑可用串口下拉控件↑↑↑↑↑↑↑↑↑

            //↓↓↓↓↓↓↓↓↓默认设置↓↓↓↓↓↓↓↓↓
            ComPort.ReadTimeout = 8000;//串口读超时8秒
            ComPort.WriteTimeout = 8000;//串口写超时8秒，在1ms自动发送数据时拔掉串口，写超时5秒后，会自动停止发送，如果无超时设定，这时程序假死
            ComPort.ReadBufferSize = 1024;//数据读缓存
            ComPort.WriteBufferSize = 1024;//数据写缓存
            //↑↑↑↑↑↑↑↑↑默认设置↑↑↑↑↑↑↑↑↑

            FileSaveTask.Elapsed += new System.Timers.ElapsedEventHandler(FileSave);//添加文件保存任务
            AskTerminalTask.Elapsed += new System.Timers.ElapsedEventHandler(AskTerminal);//添加发送终端指令任务
            CollectTimesTextBlock.Text = CollectCount.ToString();//显示采集次数

            ComPort.DataReceived += new SerialDataReceivedEventHandler(ComReceive);//串口接收中断

            Thread _ComRec = new Thread(new ThreadStart(ComRec)); //查询串口接收数据线程声明
            _ComRec.Start();//启动线程

            StartButton_Click(this, e);

        }

        private bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        /*函数名：HttpPost(string Url, string postDataStr)
         * 参数Url：要发送的服务器网址
         * 参数postDataStr：JSON格式的数据
         * 返回：服务器状态返回
         */
        private string HttpPost(string Url, string postDataStr)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);

            ServicePointManager.ServerCertificateValidationCallback = ValidateServerCertificate;

            request.Method = "POST";
            request.ContentType = "application/json;charset=utf-8";
            request.ContentLength = Encoding.UTF8.GetByteCount(postDataStr);
            Stream myRequestStream = request.GetRequestStream();
            StreamWriter myStreamWriter = new StreamWriter(myRequestStream, Encoding.GetEncoding("gb2312"));
            myStreamWriter.Write(postDataStr);
            myStreamWriter.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();

            return retString;
        }

        //发送终端指令任务实现
        private void AskTerminal(object source, ElapsedEventArgs e)
        {
            try
            {         
                ComPort.Write(Terminals, 0, 3);
                //延迟CollectIntervalTime秒
                for (int i = 0; i < CollectIntervalTime; i++)
                {
                    Thread.Sleep(1000);
                }
                ComPort.Write(Terminals, 3, 3);
                //延迟CollectIntervalTime秒
                for (int i = 0; i < CollectIntervalTime; i++)
                {
                    Thread.Sleep(1000);
                }
                ComPort.Write(Terminals, 6, 3);
            }
            catch (Exception err)
            {
                string ErrorMessage = "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "发送指令错误：" + err.Message;
                Dispatcher.Invoke(new Action(() =>
                {
                    SystemMessageTextBox.Text += ErrorMessage;
                    SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                }));
            }
        }

        //日志文件保存任务实现
        private void FileSave(object source, ElapsedEventArgs e)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                string CurrentTime, Path, JSONFileSaveReturn,LogFileSaveReturn;
                CurrentTime = DateTime.Now.ToString("yyyy年MM月dd日HH时mm分ss秒");

                //保存data数据
                Path = JSONSaveDocument + "\\" + CurrentTime + ".txt";
                JSONFileSaveReturn = SaveFile(Path, SendDataTextBox.Text);
                SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "文件保存：" + JSONFileSaveReturn;

                //保存运行信息
                Path = JSONSaveDocument + "\\" + CurrentTime + " log.txt";
                LogFileSaveReturn = SaveFile(Path, SystemMessageTextBox.Text);                 
                SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "文件保存：" + LogFileSaveReturn;
                SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面

                //清空data框和运行信息框
                SendDataTextBox.Text = "";
                SystemMessageTextBox.Text = "";
            }));
        }

        //显示当前时间
        private void CurrentTimeShow()
        {
            string StartMessage = "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "时钟线程正在运行...";
            Dispatcher.Invoke(new Action(() =>
            {
                SystemMessageTextBox.Text += StartMessage;
                SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
            }));

            while (true)
            {
                string UpdateMessage = DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss");
                Dispatcher.Invoke(new Action(() =>
                {
                    CurrentTimeTextBlock.Text = UpdateMessage;
                }));
            }
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)//打开/关闭串口事件
        {
            if (AvailableComCbobox.SelectedValue == null)//先判断是否有可用串口
            {
                SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "无可用串口，无法启动服务!";
                SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                return;//没有串口，提示后直接返回
            }

            #region 打开服务
            if (ComPortIsOpen == false)//ComPortIsOpen == false当前串口为关闭状态，按钮事件为打开串口
            {
                //↓↓↓↓↓↓↓↓↓获取参数↓↓↓↓↓↓↓↓↓ 
                Url = AddressTextBox.Text;//获取服务器网址
                PollingIntervalTime = int.Parse(PollingIntervalTimeTextBox.Text) * 1000;//获取轮询间隔时间
                AskTerminalTask.Interval = PollingIntervalTime;//设置发送指令任务执行间隔时间
                FileSavePeriod = int.Parse(FileSavePeriodTextBox.Text) * 1000;//获取文件保存间隔时间
                FileSaveTask.Interval = FileSavePeriod;//设置文件保存任务间隔时间
                CollectIntervalTime = int.Parse(CollectIntervalTimeTextBox.Text);//获取采集间隔
                //↑↑↑↑↑↑↑↑↑获取参数↑↑↑↑↑↑↑↑↑

                try//尝试打开串口
                {
                    ComPort.PortName = AvailableComCbobox.SelectedValue.ToString();//设置要打开的串口
                    ComPort.BaudRate = 9600;//设置当前波特率
                    ComPort.Parity = (Parity)0;//设置当前校验位
                    ComPort.DataBits = 8;//设置当前数据位
                    ComPort.StopBits = (StopBits)1;//设置当前停止位                    
                    ComPort.Open();//打开串口
                    SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "串口已启动!";
                }
                catch (Exception err)//如果串口被其他占用，则无法打开
                {
                    SystemMessageTextBox.Text += DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + err.Message;
                    SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                    GetPort();//刷新当前可用串口
                    return;//无法打开串口，提示后直接返回
                }

                
                //↓↓↓↓↓↓↓↓↓发送起始数据到服务器↓↓↓↓↓↓↓↓↓ 
                string TestString = "[{\"tSerialNumber\":\"343\"}]";
                try
                {
                    SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + HttpPost(Url, TestString);//显示服务器返回信息
                }
                catch (Exception err)
                {
                    SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + err.Message;
                    ComPort.Close();
                    SystemMessageTextBox.Text = "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "串口已停止!";
                    return;//无法连接服务器
                }
                //↑↑↑↑↑↑↑↑↑发送起始数据到连接服务器↑↑↑↑↑↑↑↑↑                  
                

                AskTerminalTask.Start();//启动发送终端指令任务
                FileSaveTask.Start();//启动文件保存任务

                //↓↓↓↓↓↓↓↓↓成功打开服务后的设置↓↓↓↓↓↓↓↓↓
                StartButton.Content = "关闭服务";//按钮显示改为“关闭服务”
                ComPortIsOpen = true;//串口打开状态字改为true
                WaitClose = false;//等待关闭串口状态改为false     
                AvailableComCbobox.IsEnabled = false;//失能可用串口控件
                AddressTextBox.IsEnabled = false;//失能IP输入框编辑
                PollingIntervalTimeTextBox.IsEnabled = false;//禁止编辑轮询间隔时间
                FileSavePeriodTextBox.IsEnabled = false;//禁止编辑POST GET
                CollectIntervalTimeTextBox.IsEnabled = false;//禁止编辑采集间隔时间
                StartTimeTextBlock.Text = DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss");//显示启动时间 
                SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "启动成功！";
                //↑↑↑↑↑↑↑↑↑成功打开串口后的设置↑↑↑↑↑↑↑↑↑                       
            }
            #endregion
            #region 关闭服务
            else//ComPortIsOpen == true,当前串口为打开状态，按钮事件为关闭串口
            {
                AskTerminalTask.Stop();//停止发送终端指令任务
                FileSaveTask.Stop();//停止文件保存任务

                //文件保存
                if (SendDataTextBox.Text != "")
                {
                    MessageBoxResult SaveFileResult = MessageBox.Show("是否保存Data框的数据？", "文件保存", MessageBoxButton.YesNo);//显示确认窗口
                    if (SaveFileResult == MessageBoxResult.Yes)
                    {
                        string nowtime, Path, SaveFileReturnMessage;

                        //保存JSON数据
                        nowtime = DateTime.Now.ToString("yyyy年MM月dd日HH时mm分ss秒");
                        Path = JSONSaveDocument + "\\" + nowtime + ".txt";
                        SaveFileReturnMessage = SaveFile(Path, SendDataTextBox.Text);
                        MessageBox.Show("文件保存：" + SaveFileReturnMessage);

                        //保存运行日志
                        Path = JSONSaveDocument + "\\" + nowtime + " log.txt";
                        SaveFile(Path, SystemMessageTextBox.Text);

                        MessageBox.Show("文件保存：" + Path);
                    }
                }

                //尝试关闭串口
                try
                {
                    ComPort.DiscardOutBuffer();//清发送缓存
                    ComPort.DiscardInBuffer();//清接收缓存
                    WaitClose = true;//激活正在关闭状态字，用于在串口接收方法的invoke里判断是否正在关闭串口
                    while (Listening)//判断invoke是否结束
                    {
                        DispatcherHelper.DoEvents(); //循环时，仍进行等待事件中的进程，该方法为winform中的方法，WPF里面没有，这里在后面自己实现
                    }
                    ComPort.Close();//关闭串口
                    WaitClose = false;//关闭正在关闭状态字，用于在串口接收方法的invoke里判断是否正在关闭串口
                    SetAfterClose();//成功关闭串口或串口丢失后的设置
                    AvailableComCbobox.IsEnabled = true;//使能可用串口控件
                    FileSavePeriodTextBox.IsEnabled = true;//使能端口修改                                    
                }
                catch (Exception err)//如果在未关闭串口前，串口就已丢失，这时关闭串口会出现异常
                {
                    SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "关闭串口失败：" + err.Message;
                    SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                    if (ComPort.IsOpen == false)//判断当前串口状态，如果ComPort.IsOpen==false，说明串口已丢失
                    {
                        SetComLose();
                    }
                    else//无法关闭串口
                    {
                        SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "无法关闭串口，原因未知！";
                        SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                        return;//无法关闭串口，提示后直接返回
                    }
                }              

                //使能相关控件
                AddressTextBox.IsEnabled = true;//使能IP输入框编辑
                PollingIntervalTimeTextBox.IsEnabled = true;//使能轮询间隔时间修改
                CollectIntervalTimeTextBox.IsEnabled = true;//使能采集间隔时间修改  
                CollectCount = 0;//清空采集次数
            }
            #endregion
        }

        private void AvailableComCbobox_PreviewMouseDown(object sender, MouseButtonEventArgs e)//刷新可用串口
        {
            GetPort();//刷新可用串口
        }

        private void GetPort()//刷新可用串口
        {
            comList.Clear();//清空控件链接资源
            AvailableComCbobox.DisplayMemberPath = "com1";
            AvailableComCbobox.SelectedValuePath = null;//路径都指为空，清空下拉控件显示，下面重新添加

            ports = new string[SerialPort.GetPortNames().Length];//重新定义可用串口数组长度
            ports = SerialPort.GetPortNames();//获取可用串口
            if (ports.Length > 0)//有可用串口
            {
                for (int i = 0; i < ports.Length; i++)
                {
                    comList.Add(new customer() { com = ports[i] });//下拉控件里添加可用串口
                }
                AvailableComCbobox.ItemsSource = comList;//可用串口下拉控件资源路径
                AvailableComCbobox.DisplayMemberPath = "com";//可用串口下拉控件显示路径
                AvailableComCbobox.SelectedValuePath = "com";//可用串口下拉控件值路径
            }
        }

        public class customer//可用串口号访问接口
        {
            public string com { get; set; }//可用串口
        }

        private void ComReceive(object sender, SerialDataReceivedEventArgs e)//接收数据 中断只标志有数据需要读取，读取操作在中断外进行
        {
            if (WaitClose) return;//如果正在关闭串口，则直接返回
            Thread.Sleep(10);//发送和接收均为文本时，接收中为加入判断是否为文字的算法，发送你（C4E3），接收可能识别为C4,E3，可用在这里加延时解决
            byte[] recBuffer;//接收缓冲区
            try
            {
                System.Threading.Thread.Sleep(300);//等待47个字节全部接收完毕，目的是为了兼容计算机和终端设备的处理速度
                recBuffer = new byte[ComPort.BytesToRead];//接收数据缓存大小
                ComPort.Read(recBuffer, 0, recBuffer.Length);//读取数据
                recQueue.Enqueue(recBuffer);//读取数据入列Enqueue（全局）
            }
            catch (Exception err)
            {
                UIAction(() =>
                {
                    if (ComPort.IsOpen == false)//如果ComPort.IsOpen == false，说明串口已丢失
                    {
                        SetComLose();//串口丢失后相关设置
                    }
                    else
                    {
                        SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "接收数据错误：" + err.Message;
                        SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                    }
                });
            }
        }

        void ComRec()//数据处理线程，窗口初始化中就开始启动运行，对串口接收到的数据进行处理
        {
            while (true)//一直查询串口接收线程中是否有新数据
            {
                if (recQueue.Count > 0)//当串口接收线程中有新的数据时候，队列中有新进的成员recQueue.Count > 0
                {
                    byte[] recBuffer = (byte[])recQueue.Dequeue();//出列Dequeue（全局）
                    if (recBuffer.Length > 44)//判断是否终端发送的字节
                    {
                        try//尝试进行数据赋值操作
                        {
                            //设备编号
                            TerminalNumber = recBuffer[0];
                            //空气湿度
                            AtmosphereHumidity_H = recBuffer[1];
                            AtmosphereHumidity_L = recBuffer[2];
                            AtmosphereHumidity = (UInt16)((AtmosphereHumidity_H << 8) | AtmosphereHumidity_L);
                            //空气温度
                            AtmosphereTemperature_H = recBuffer[3];
                            AtmosphereTemperature_L = recBuffer[4];
                            AtmosphereTemperature = (UInt16)((AtmosphereTemperature_H << 8) | AtmosphereTemperature_L);
                            //光照强度
                            Illumination_01 = recBuffer[5];
                            Illumination_02 = recBuffer[6];
                            Illumination_03 = recBuffer[7];
                            Illumination_04 = recBuffer[8];
                            Illumination = (UInt32)((Illumination_01 << 24) | (Illumination_02 << 16) | (Illumination_03 << 8) | Illumination_04);
                            //01号土壤温湿度
                            SoilHumidity_H = recBuffer[9];
                            SoilHumidity_L = recBuffer[10];
                            SoilHumidity01 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[11];
                            SoilTemperature_L = recBuffer[12];
                            SoilTemperature01 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //02号土壤温湿度
                            SoilHumidity_H = recBuffer[13];
                            SoilHumidity_L = recBuffer[14];
                            SoilHumidity02 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[15];
                            SoilTemperature_L = recBuffer[16];
                            SoilTemperature02 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //04号土壤温湿度
                            SoilHumidity_H = recBuffer[17];
                            SoilHumidity_L = recBuffer[18];
                            SoilHumidity03 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[19];
                            SoilTemperature_L = recBuffer[20];
                            SoilTemperature03 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //04号土壤温湿度
                            SoilHumidity_H = recBuffer[21];
                            SoilHumidity_L = recBuffer[22];
                            SoilHumidity04 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[23];
                            SoilTemperature_L = recBuffer[24];
                            SoilTemperature04 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //05号土壤温湿度
                            SoilHumidity_H = recBuffer[25];
                            SoilHumidity_L = recBuffer[26];
                            SoilHumidity05 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[27];
                            SoilTemperature_L = recBuffer[28];
                            SoilTemperature05 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //06号土壤温湿度
                            SoilHumidity_H = recBuffer[29];
                            SoilHumidity_L = recBuffer[30];
                            SoilHumidity06 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[31];
                            SoilTemperature_L = recBuffer[32];
                            SoilTemperature06 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //07号土壤温湿度
                            SoilHumidity_H = recBuffer[33];
                            SoilHumidity_L = recBuffer[34];
                            SoilHumidity07 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[35];
                            SoilTemperature_L = recBuffer[36];
                            SoilTemperature07 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //08号土壤温湿度
                            SoilHumidity_H = recBuffer[37];
                            SoilHumidity_L = recBuffer[38];
                            SoilHumidity08 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[39];
                            SoilTemperature_L = recBuffer[40];
                            SoilTemperature08 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //09号土壤温湿度
                            SoilHumidity_H = recBuffer[41];
                            SoilHumidity_L = recBuffer[42];
                            SoilHumidity09 = (UInt16)((SoilHumidity_H << 8) | SoilHumidity_L);
                            SoilTemperature_H = recBuffer[43];
                            SoilTemperature_L = recBuffer[44];
                            SoilTemperature09 = (UInt16)((SoilTemperature_H << 8) | SoilTemperature_L);
                            //土壤酸碱度传感器
                            SoliPH_H = recBuffer[45];
                            SOliPH_L = recBuffer[46];
                            SoliPH = (UInt16)((SoliPH_H << 8) | SOliPH_L);

                            //建立JSON字符串(一次发送格式)                       
                            StringBuilder JsonMessage = new StringBuilder();
                            JsonMessage.Append("[");
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}001\",\"tTemperature\":{1}.{2},\"tHumidity\":{3}.{4},\"tLightIntensity\":{5}}},",
                                TerminalNumber, AtmosphereTemperature / 10, AtmosphereTemperature % 10, AtmosphereHumidity / 10, AtmosphereHumidity % 10, Illumination);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}002\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature01 / 10, SoilTemperature01 % 10, SoilHumidity01 / 10, SoilHumidity01 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}003\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature02 / 10, SoilTemperature02 % 10, SoilHumidity02 / 10, SoilHumidity02 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}004\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature03 / 10, SoilTemperature03 % 10, SoilHumidity03 / 10, SoilHumidity03 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}005\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature04 / 10, SoilTemperature04 % 10, SoilHumidity04 / 10, SoilHumidity04 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}006\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature05 / 10, SoilTemperature05 % 10, SoilHumidity05 / 10, SoilHumidity05 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}007\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature06 / 10, SoilTemperature06 % 10, SoilHumidity06 / 10, SoilHumidity06 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}008\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature07 / 10, SoilTemperature07 % 10, SoilHumidity07 / 10, SoilHumidity07 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}009\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature08 / 10, SoilTemperature08 % 10, SoilHumidity08 / 10, SoilHumidity08 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}010\",\"tSoilTemperature\":{1}.{2},\"tSoilHumidity\":{3}.{4}}},",
                                TerminalNumber, SoilTemperature09 / 10, SoilTemperature09 % 10, SoilHumidity09 / 10, SoilHumidity09 % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.AppendFormat("{{\"tSerialNumber\":\"{0:D2}011\",\"tSoilPh\":{1}.{2}}}",
                                TerminalNumber, SoliPH / 10, SoliPH % 10);//注意要输入两个大括号，来显示一个大括号
                            JsonMessage.Append("]");

                            string JSONData = JsonMessage.ToString();//转码

                            CollectCount++;//采集次数加1

                            //显示数据到面板中
                            UIAction(() =>
                            {
                                CollectTimesTextBlock.Text = CollectCount.ToString();//显示采集次数
                                SendDataTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss");//显示时间
                                SendDataTextBox.Text += "\r\n" + JSONData + "\r\n";//显示即将发送到服务器的数据
                                SendDataTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                            });
                            //尝试发送数据到服务器
                            try
                            {
                                HttpPost(Url, JSONData);
                            }
                            catch (Exception err)//如果发送服务器失败
                            {
                                string ErrorMessage = err.Message;//异常信息内容
                                UIAction(() =>
                                {
                                    SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "上传服务器错误：" + ErrorMessage;
                                    SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                                });
                            }
                        }
                        catch (Exception err)//如果赋值失败
                        {
                            Dispatcher.Invoke(new Action(() =>
                            {
                                SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "数据赋值错误：" + err.Message;
                                SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
                            }));
                        }
                    }
                }
                else
                {
                    Thread.Sleep(100);//如果不延时，一直查询，将占用CPU过高
                }
            }
        }

        private void SetComLose()//成功关闭串口或串口丢失后的设置
        {
            WaitClose = true;//;//激活正在关闭状态字，用于在串口接收方法的invoke里判断是否正在关闭串口
            while (Listening)//判断invoke是否结束
            {
                DispatcherHelper.DoEvents(); //循环时，仍进行等待事件中的进程，该方法为winform中的方法，WPF里面没有，这里在后面自己实现
            }
            Dispatcher.Invoke(new Action(() =>
            {
                SystemMessageTextBox.Text += "\r\n" + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss :") + "串口已丢失！";
                SystemMessageTextBoxScrollViewer.ScrollToBottom();//滚到最下面
            }));
            WaitClose = false;//关闭正在关闭状态字，用于在串口接收方法的invoke里判断是否正在关闭串口
            GetPort();//刷新可用串口
            SetAfterClose();
        }

        private void SetAfterClose()//成功关闭串口或串口丢失后的设置
        {
            StartButton.Content = "启动服务";//按钮显示为“启动服务”
            ComPortIsOpen = false;//串口状态设置为关闭状态
            AvailableComCbobox.IsEnabled = true;//使能可用串口控件
            StartTimeTextBlock.Text = "";
            CollectTimesTextBlock.Text = "";
            SendDataTextBox.Text = "";
            SystemMessageTextBox.Text = "";
        }

        //模拟 Winfrom 中 Application.DoEvents() 详见 http://www.silverlightchina.net/html/study/WPF/2010/1216/4186.html?1292685167
        public static class DispatcherHelper
        {
            [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
            public static void DoEvents()
            {
                DispatcherFrame frame = new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrames), frame);
                try { Dispatcher.PushFrame(frame); }
                catch (InvalidOperationException) { }
            }
            private static object ExitFrames(object frame)
            {
                ((DispatcherFrame)frame).Continue = false;
                return null;
            }
        }

        void UIAction(Action action)//在主线程外激活线程方法
        {
            System.Threading.SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext(App.Current.Dispatcher));
            System.Threading.SynchronizationContext.Current.Post(_ => action(), null);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)//关闭窗口closing
        {
            MessageBoxResult QuitResult = MessageBox.Show("确认是否要退出？", "退出", MessageBoxButton.YesNo);//显示确认窗口
            if (QuitResult == MessageBoxResult.Yes)
            {
                if (SendDataTextBox.Text != "")
                {
                    MessageBoxResult SaveFileResult = MessageBox.Show("退出前是否保存Data框的数据？", "文件保存", MessageBoxButton.YesNo);//显示确认窗口
                    if (SaveFileResult == MessageBoxResult.Yes)
                    {                        
                        string nowtime, Path, SaveFileReturnMessage;

                        //保存JSON数据
                        nowtime = DateTime.Now.ToString("yyyy年MM月dd日HH时mm分ss秒");
                        Path = JSONSaveDocument + "\\" + nowtime + ".txt";
                        SaveFileReturnMessage = SaveFile(Path, SendDataTextBox.Text);
                        MessageBox.Show("文件保存：" + SaveFileReturnMessage);

                        //保存运行日志
                        Path = JSONSaveDocument + "\\" + nowtime + " log.txt";
                        SaveFile(Path, SystemMessageTextBox.Text);
                    }
                }
            }
            else
            {
                e.Cancel = true;//取消操作
            }
        }

        private void Window_Closed(object sender, EventArgs e)//关闭窗口确认后closed ALT+F4
        {
            Application.Current.Shutdown();//先停止线程,然后终止进程.
            Environment.Exit(0);//直接终止进程.
        }

        /**
         *功能：数据保存到文件中
         *参数：
         *  Path：要保存到的目的路径
         *  Data：要保存的数据
         *返回：
         *  保存成功则返回目的路径，失败则返回错误信息。
         */
        private string SaveFile(string Path, string Data)
        {
            try
            {
                if (!(Data == ""))
                {
                    if (!Directory.Exists(JSONSaveDocument))
                    {
                        Directory.CreateDirectory(JSONSaveDocument);
                    }
                    FileStream fs = new FileStream(Path, FileMode.Create);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.Write(Data);
                    sw.Close();
                    fs.Close();
                    return Path;
                }
                else
                {
                    return "数据为空，不保存。";
                }
            }
            catch (Exception err)
            {
                return "保存失败！" + err.Message;
            }
        }
    }
}