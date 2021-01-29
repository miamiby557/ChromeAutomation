using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using WebSocketSharp;
using WebSocketSharp.Server;

namespace ChromeAutomation
{
    public struct CursorPoint
    {
        public int X;

        public int Y;
    }
    partial class Form1
    {

        [DllImport("gdi32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int CreatePen(int nPenStyle, int nWidth, int crColor);

        [DllImport("gdi32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int SelectObject(int hdc, int hObject);

        [DllImport("gdi32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int SetROP2(int hdc, int nDrawMode);

        [DllImport("gdi32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int Rectangle(int hdc, int X1, int Y1, int X2, int Y2);

        [DllImport("user32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int GetDC(int hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int ReleaseDC(int hwnd, int hdc);

        [DllImport("gdi32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int DeleteObject(int hObject);

        [DllImport("user32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool GetPhysicalCursorPos(ref ChromeAutomation.CursorPoint lpPoint);

        public delegate void RS_DATA_CALLBACK(int left, int top, int width, int height, Dictionary<string, object> dictionary);

        private static RS_DATA_CALLBACK MessageCallback = null;

        public static Form1.SocketBehavior CureentSocketBehavior = null;

        public static void RegisterMessageCallback(RS_DATA_CALLBACK callback)
        {
            Form1.MessageCallback = callback;
        }

        public class SocketBehavior : WebSocketBehavior {
            protected override void OnOpen()
            {
                base.OnOpen();
                Console.WriteLine("WebSocket OnOpen");
                Form1.CureentSocketBehavior = this;
                Send("{\"event\":\"UI\"}");
            }

            

            protected override void OnMessage(MessageEventArgs e)
            {
                // Console.WriteLine("WebSocket OnMessage:" + e.Data);
                JObject jo = (JObject)JsonConvert.DeserializeObject(e.Data);
                if (jo.ContainsKey("event") && jo["event"].ToString().Equals("UI") && jo.ContainsKey("data"))
                {
                    Console.WriteLine("来自客户端的xpath消息:" + e.Data.ToString());
                    try
                    {
                        IHTML iHTML = JsonConvert.DeserializeObject<IHTML>(jo["data"].ToString());
                        int left = iHTML.left;
                        int top = iHTML.top;
                        int width = iHTML.width;
                        int height = iHTML.height;
                        Dictionary<string, object> dictionary = new Dictionary<string, object>();
                        dictionary.Add("id", iHTML.id);
                        dictionary.Add("tag", iHTML.tag);
                        dictionary.Add("text", iHTML.text);
                        dictionary.Add("value", iHTML.value);
                        dictionary.Add("xpath", iHTML.xpath);
                        dictionary.Add("full_xpath", iHTML.full_xpath);
                        MessageCallback(left, top, width, height, dictionary);
                    }
                    catch (Exception)
                    {
                    }
                    bool flag2 = ((int)Form1.GetAsyncKeyState(17) & 32768) != 0;
                    if (flag2)
                    {
                        Application.Exit();
                    }
                    
                }
                else if(jo.ContainsKey("event") && jo["event"].ToString().Equals("CHECK"))
                {
                    /*JObject position = (JObject)JsonConvert.DeserializeObject(jo["position"].ToString());
                    Dictionary<string, object> dictionary = new Dictionary<string, object>();
                    dictionary.Add("event", "UI");
                    dictionary.Add("data", new Dictionary<string, object>
                            {
                                {
                                    "x",
                                    position["x"]
                                },
                                {
                                    "y",
                                    position["y"]
                                }
                            });
                    // Console.WriteLine("发送消息给客户端:" + JsonConvert.SerializeObject(dictionary));
                    SendMessage(JsonConvert.SerializeObject(dictionary));*/
                    System.Drawing.Point point = new System.Drawing.Point(-1, -1);
                    ChromeAutomation.CursorPoint cursorPoint = default(ChromeAutomation.CursorPoint);
                    Form1.GetPhysicalCursorPos(ref cursorPoint);
                    System.Drawing.Point point2 = new System.Drawing.Point(cursorPoint.X, cursorPoint.Y);
                    Console.WriteLine("cursorPoint：x:" + cursorPoint.X.ToString() + ",y:" + cursorPoint.Y.ToString());
                    bool flag4 = CureentSocketBehavior != null;
                    Console.WriteLine("Form1.CureentSocketBehavior:" + flag4);
                    if (flag4)
                    {
                        int x = point2.X;
                        int y = point2.Y;
                        Dictionary<string, object> dictionary = new Dictionary<string, object>();
                        dictionary.Add("event", "UI");
                        dictionary.Add("data", new Dictionary<string, object>
                            {
                                {
                                    "x",
                                    x
                                },
                                {
                                    "y",
                                    y
                                }
                            });
                        Console.WriteLine("鼠标坐标：x:" + x.ToString() + ",y:" + y.ToString());
                        SendMessage(JsonConvert.SerializeObject(dictionary));
                    }
                }
            }

            public void SendMessage(string msg)
            {
                base.Send(msg);
            }
        }

        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private int hDC;

        private int left = 0, top = 0, width = 0, height = 0;

        Dictionary<string, object> dictionary;

        public void DrawRect(int left, int top, int width, int height, Dictionary<string, object> dictionary)
        {
            /*Graphics graphics = this.CreateGraphics();
            Pen mypen = new Pen(Color.Red, 3);//设置画笔属性
            graphics.DrawRectangle(mypen, 100, 100, 50, 50);//按按钮画一个矩形*/
            // 回调实现方法
            Console.WriteLine("DrawRect:" + (left, top, width, height));
            /*base.Location = new System.Drawing.Point(left, top);
            base.Size = new System.Drawing.Size(width, height);
            base.Visible = true;*/
            /*Form1.Rectangle(this.hDC, left, top, width, height);*/
            /*Form1.Rectangle(this.hDC, 200, 300, 400, 250);*/
            // (964, 266, 53, 53)
            /*Form1.Rectangle(this.hDC, 964, 266+53, 964+53, 266);*/
            this.left = left;
            this.top = top;
            this.width = width;
            this.height = height;
            this.dictionary = dictionary;
        }

        [DebuggerBrowsable(DebuggerBrowsableState.Never), AccessedThroughProperty("Timer1"), CompilerGenerated]
        private System.Windows.Forms.Timer _Timer1;

        internal virtual System.Windows.Forms.Timer Timer1
        {
            [CompilerGenerated]
            get
            {
                return this._Timer1;
            }
            [CompilerGenerated]
            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                EventHandler value2 = new EventHandler(this.Timer1_Tick);
                System.Windows.Forms.Timer timer = this._Timer1;
                if (timer != null)
                {
                    timer.Tick -= value2;
                }
                this._Timer1 = value;
                timer = this._Timer1;
                if (timer != null)
                {
                    timer.Tick += value2;
                }
            }
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            if (left > 0 && top > 0)
            {
                Form1.Rectangle(this.hDC, left, top + height, left + width, top);
                Thread.Sleep(100);
                Form1.Rectangle(this.hDC, left, top + height, left + width, top);
                Thread.Sleep(500);
            }
        }


        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        /// 

        private int Pen;

        private int PreviousPen;

        [DllImport("user32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int RegisterHotKey(int hwnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern int UnregisterHotKey(int hwnd, int id);

        [DllImport("user32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        public static extern short GetAsyncKeyState(int vKey);

        private KeyEventHandler myKeyEventHandeler = null;//按键钩子
        private KeyboardHook k_hook = new KeyboardHook();

        private void hook_KeyDown(object sender, KeyEventArgs e)
        {
            //  这里写具体实现
            Console.WriteLine("按下按键" + e.KeyValue);
            if (e.KeyValue == 162)
            {
                try
                {
                    string url = "http://localhost:63361/postChromeData";
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                    req.Method = "POST";
                    req.Timeout = 4000;//设置请求超时时间，单位为毫秒
                    req.ContentType = "application/json";
                    byte[] data = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(this.dictionary));
                    req.ContentLength = data.Length;
                    using (Stream reqStream = req.GetRequestStream())
                    {
                        reqStream.Write(data, 0, data.Length);
                        reqStream.Close();
                        // 关闭
                        this.Close();
                        Application.Exit();
                    }
                }
                catch (Exception)
                {
                    // 关闭
                    this.Close();
                    Application.Exit();
                }
            }
        }

        public void stopListen()
        {
            if (myKeyEventHandeler != null)
            {
                k_hook.KeyDownEvent -= myKeyEventHandeler;//取消按键事件
                myKeyEventHandeler = null;
                k_hook.Stop();//关闭键盘钩子
            }
        }

        private Thread selectThread = null;


        private void RealMousePosition()
        {
            System.Drawing.Point point = new System.Drawing.Point(-1, -1);
            while (true)
            {
                ChromeAutomation.CursorPoint cursorPoint = default(ChromeAutomation.CursorPoint);
                Form1.GetPhysicalCursorPos(ref cursorPoint);
                System.Drawing.Point point2 = new System.Drawing.Point(cursorPoint.X, cursorPoint.Y);
                Console.WriteLine("cursorPoint：x:" + cursorPoint.X.ToString() + ",y:" + cursorPoint.Y.ToString());
                bool flag4 = CureentSocketBehavior != null;
                Console.WriteLine("Form1.CureentSocketBehavior:" + flag4);
                if (flag4)
                {
                    int x = point2.X;
                    int y = point2.Y;
                    Dictionary<string, object> dictionary = new Dictionary<string, object>();
                    dictionary.Add("event", "UI");
                    dictionary.Add("data", new Dictionary<string, object>
                            {
                                {
                                    "x",
                                    x
                                },
                                {
                                    "y",
                                    y
                                }
                            });
                    Console.WriteLine("鼠标坐标：x:" + x.ToString() + ",y:" + y.ToString());
                    Form1.CureentSocketBehavior.SendMessage(JsonConvert.SerializeObject(dictionary));
                }
                Thread.Sleep(1000);
            }
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "谷歌浏览器拾取插件";
            CheckForIllegalCrossThreadCalls = false;
            this.hDC = Form1.GetDC(0);
            this.Pen = Form1.CreatePen(0, 3, 255);
            this.PreviousPen = Form1.SelectObject(this.hDC, this.Pen);
            Form1.SetROP2(this.hDC, 10);
            this.Timer1 = new System.Windows.Forms.Timer(this.components);
            this.Timer1.Start();
            base.Name = "CINDATA AUTOMATION";
            base.ClientSize = new System.Drawing.Size(200, 30);
            base.StartPosition = FormStartPosition.Manual;
            Rectangle rect = Screen.GetWorkingArea(this);
            System.Drawing.Point p = new System.Drawing.Point(rect.Width - 250, rect.Height - 40);
            this.Location = p;
            this.TopMost = true;
            this.ControlBox = false;

            this.selectThread = new Thread(new ThreadStart(this.RealMousePosition));
            this.selectThread.IsBackground = true;
            this.selectThread.Start();
            // websocket
            WebSocketServer webSocketServer = new WebSocketServer("ws://127.0.0.1:63360");
            string uri = "/chrome";
            webSocketServer.AddWebSocketService<SocketBehavior>(uri);
            webSocketServer.Start();
            RegisterMessageCallback(new RS_DATA_CALLBACK(this.DrawRect));

            myKeyEventHandeler = new KeyEventHandler(hook_KeyDown);
            k_hook.KeyDownEvent += myKeyEventHandeler;//钩住键按下
            k_hook.Start();//安装键盘钩子
        }

        #endregion
    }
}

