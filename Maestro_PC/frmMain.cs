using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Threading;
using System.Net.Sockets;
using InTheHand.Net.Sockets;
using InTheHand.Net.Bluetooth;
using InTheHand.Windows.Forms;
using InTheHand.Net.Bluetooth.AttributeIds;
using System.IO;
namespace Maestro_PC
{

    public partial class frmMain : MetroForm
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        static extern bool SetCursorPos(int x, int y);
        //Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        public void DoMouseClick()
        {
            //Call the imported function with the cursor's current position
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }
        private PPT.Application pptApplication;
        private PPT.Presentation presentation;
        private PPT.Slide slide;
        private PPT.Slides slides;
        private int slidescount;
        private bool direction = true;
        private bool powercon = false;
        private bool threadcon = true;
        char[] delimiterChars = { 'M', ','};
        //private double v0_X, v0_Y;
        //private double dt;
        Thread AcceptAndListeningThread;

        // helper variable
        Boolean isConnected = false;

        //bluetooth stuff
        BluetoothClient btClient;  //represent the bluetooth client connection
        BluetoothListener btListener; //represent this server bluetooth device

        MetroFramework.Controls.MetroButton button_Connect = new MetroFramework.Controls.MetroButton();
        MetroFramework.Controls.MetroButton button_Bluetooth = new MetroFramework.Controls.MetroButton();
        public frmMain()
        {
            InitializeComponent();
        }
        private void button_Connect_Click(object sender, EventArgs e)
        {
            try
            {
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPT.Application;
                presentation = pptApplication.ActivePresentation;
                slides = presentation.Slides;
                slidescount = slides.Count;
                try
                {
                    // Get selected slide object in normal view 
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // Get selected slide object in reading view 
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
                powercon = true;
                button_Connect.Text = "Powerpoint Connected";
                button_Connect.Enabled = false;
            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "Please check PowerPoint is running.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void slide_Previous()
        {
            if (powercon)
            {
                try
                {
                    pptApplication.Active.ToString();
                    int slideIndex = slide.SlideIndex - 1;
                    if (slideIndex >= 1)
                    {
                        try
                        {
                            slide = slides[slideIndex];
                            slides[slideIndex].Select();
                        }
                        catch
                        {
                            pptApplication.SlideShowWindows[1].View.Previous();
                            slide = pptApplication.SlideShowWindows[1].View.Slide;
                        }
                    }
                    else
                    {
                        //MetroFramework.MetroMessageBox.Show(this, "This page is the first page.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch
                {
                    powercon = false;
                    button_Connect.Text = "Powerpoint Connection";
                    button_Connect.Enabled = true;
                    MetroFramework.MetroMessageBox.Show(this, "Powerpoint is disconnected.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }
        private void button_Bluetooth_Click(object sender, EventArgs e)
        {
            bluetooth_connect();
        }
        private void slide_next()
        {
            if (powercon)
            {
                try
                {
                    pptApplication.Active.ToString();
                    int slideIndex = slide.SlideIndex + 1;
                    if (slideIndex > slidescount)
                    {
                        //presentation.SlideShowSettings.Application.Quit();
                        //MetroFramework.MetroMessageBox.Show(this, "This page is the last page.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        try
                        {
                            slide = slides[slideIndex];
                            slides[slideIndex].Select();
                        }
                        catch
                        {
                            pptApplication.SlideShowWindows[1].View.Next();
                            slide = pptApplication.SlideShowWindows[1].View.Slide;
                        }
                    }
                }
                catch
                {
                    powercon = false;
                    button_Connect.Text = "Powerpoint Connection";
                    button_Connect.Enabled = true;
                    MetroFramework.MetroMessageBox.Show(this, "Powerpoint is disconnected.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void bluetooth_connect()
        {
            if (BluetoothRadio.IsSupported)
            {

                UpdateLogText("Bluetooth Supported!");
                UpdateLogText("—————————–");

                //getting device information
                UpdateLogText("Primary Bluetooth Radio Name : " + BluetoothRadio.PrimaryRadio.Name);
                UpdateLogText("Primary Bluetooth Radio Address : " + BluetoothRadio.PrimaryRadio.LocalAddress);
                UpdateLogText("Primary Bluetooth Radio Manufacturer : " + BluetoothRadio.PrimaryRadio.Manufacturer);
                UpdateLogText("Primary Bluetooth R  adio Mode : " + BluetoothRadio.PrimaryRadio.Mode);
                UpdateLogText("Primary Bluetooth Radio Software Manufacturer : " + BluetoothRadio.PrimaryRadio.SoftwareManufacturer);
                UpdateLogText("—————————–");

                button_Bluetooth.Enabled = false;
                button_Bluetooth.Text = "Waiting for watch connection";
                threadcon = true;
                //creating and starting the thread
                AcceptAndListeningThread = new Thread(AcceptAndListen);
                AcceptAndListeningThread.IsBackground = true;
                AcceptAndListeningThread.Start();
            }
            else
            {
                UpdateLogText("Bluetooth not Supported!");
                threadcon = false;
                MetroFramework.MetroMessageBox.Show(this, "Bluetooth is not supported. Please check your computer's bluetooth turend on.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            this.Size = new Size(300, 300);

            button_Connect.Size = new Size(200, 50);
            button_Connect.Left = 50;
            button_Connect.Top = 100;
            button_Connect.Text = "Connect Powerpoint";
            button_Connect.Click += new EventHandler(button_Connect_Click);
            this.Controls.Add(button_Connect);

            button_Bluetooth.Size = new Size(200, 50);
            button_Bluetooth.Left = 50;
            button_Bluetooth.Top = 170;
            button_Bluetooth.Text = "Connect Bluetooth";
            button_Bluetooth.Click += new EventHandler(button_Bluetooth_Click);
            this.Controls.Add(button_Bluetooth);

            cbxReverse.Location = new Point(50, 250);

        }
        private delegate void UpdateLogCallback(string strMessage);
        StreamReader srReceiver;
        private void ReceiveMessages()
        {
            srReceiver = new StreamReader(btClient.GetStream());
            while (threadcon)
            {
                this.Invoke(new UpdateLogCallback(this.UpdateLogText), new object[] { srReceiver.ReadLine() });
            }
            //MetroFramework.MetroMessageBox.Show(this, "Bluetooth is disconnected.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
        public void AcceptAndListen()
        {
            while (threadcon)
            {
                if (isConnected)
                {
                    try
                    {
                        this.button_Bluetooth.Invoke((MethodInvoker)delegate { this.button_Bluetooth.Text = "Bluetooth Connected"; });
                        UpdateLogTextFromThread("Listening….");
                        ReceiveMessages();
                    }
                    catch (Exception ex)
                    {
                        UpdateLogTextFromThread("There is an error while listening connection");
                        UpdateLogTextFromThread(ex.Message);
                        isConnected = btClient.Connected;
                    }
                }
                else
                {
                    try
                    {
                        btListener = new BluetoothListener(BluetoothService.RFCommProtocol);

                        UpdateLogTextFromThread("Listener created with TCP Protocol service " + BluetoothService.RFCommProtocol);
                        UpdateLogTextFromThread("*Starting Listener….");
                        btListener.Start();
                        UpdateLogTextFromThread("Listener Started!");
                        UpdateLogTextFromThread("Accepting incoming connection….");
                        btClient = btListener.AcceptBluetoothClient();
                        isConnected = btClient.Connected;
                        UpdateLogTextFromThread("A Bluetooth Device Connected!");
                    }
                    catch (Exception e)
                    {
                        UpdateLogTextFromThread("There is an error while accepting connection");
                        UpdateLogTextFromThread(e.Message);
                        UpdateLogTextFromThread("Retrying….");
                    }
                }
            }
        }
        public void mouse_control(String msg)
        {
            int sens = 7;
            string[] words = msg.Split(delimiterChars);
            double ax = Convert.ToDouble(words[1]);
            double ay = Convert.ToDouble(words[2]);
            int x, y;
            x = (int)ax;
            y = (int)ay;
            x *= -sens;
            y *= -sens;
            Cursor.Position = new Point(Cursor.Position.X + x, Cursor.Position.Y + y);
        }
        public void UpdateLogText(String msg)
        {
            Console.WriteLine(msg);
            if (msg != null)
            { 
                if (msg[0] == '1')
                {
                    if (direction)
                        slide_Previous();
                    else
                        slide_next();
                }
                else if (msg[0] == '2')
                {
                    if (direction)
                        slide_next();
                    else
                        slide_Previous();
                }
                else if (msg[0] == '3')
                {
                    if (powercon)
                    {
                        try
                        { 
                            presentation.SlideShowWindow.View.LaserPointerEnabled = false;
                            presentation.SlideShowWindow.View.PointerType = PPT.PpSlideShowPointerType.ppSlideShowPointerArrow;
                        }
                        catch
                        { }
                    }
                }
                else if (msg[0] == '4')
                {
                    if (powercon)
                    {
                        try
                        {
                            presentation.SlideShowWindow.View.LaserPointerEnabled = true;
                            presentation.SlideShowWindow.View.PointerType = PPT.PpSlideShowPointerType.ppSlideShowPointerArrow;
                        }
                        catch { }
                    }
                }
                //else if (msg[0] == '5')
                //{
                //    if (powercon)
                //    {
                //        try
                //        { 
                //            presentation.SlideShowWindow.View.LaserPointerEnabled = false;
                //            presentation.SlideShowWindow.View.PointerType = PPT.PpSlideShowPointerType.ppSlideShowPointerPen;
                //        }
                //        catch { }
                //    }
                //}
                else if (msg[0] == 'M')
                {
                    mouse_control(msg);
                }
                else if(msg[0] == 'D')
                {
                    threadcon = false;
                    button_Bluetooth.Enabled = true;
                    button_Bluetooth.Text = "Connect Bluetooth";
                    btClient.GetStream().Close();
                    btClient.Dispose();
                    btListener.Stop();
                    AcceptAndListeningThread.Abort();
                }
                else if(msg[0]=='C')
                {
                    //click
                    DoMouseClick();
                }
                else if(msg[0]=='S')
                {
                    if(powercon)
                    {
                        try
                        {
                            presentation.SlideShowWindow.ToString();
                        }
                        catch
                        {
                            presentation.SlideShowSettings.Run();
                        }
                    }
                }
                else if(msg[0] == 'Z')
                {
                    Console.WriteLine("Z is inputed");
                    presentation.SlideShowWindow.View.PointerType = PPT.PpSlideShowPointerType.ppSlideShowPointerPen;
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                }
                else if(msg[0] == 'X')
                {
                    Console.WriteLine("X is inputed");
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    presentation.SlideShowWindow.View.PointerType = PPT.PpSlideShowPointerType.ppSlideShowPointerArrow;
                }
            }
        }
        delegate void UpdateLogTextFromThreadDelegate(String msg);
        public void UpdateLogTextFromThread(String msg)
        {
            if (!this.IsDisposed && this.InvokeRequired && threadcon)
            {
                this.Invoke(new UpdateLogTextFromThreadDelegate(UpdateLogText), new Object[] { msg });
            }
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                threadcon = false;
                btListener.Stop();
                if(btClient != null)
                {
                    btClient.GetStream().Close();
                    btClient.Dispose();
                }
                AcceptAndListeningThread.Abort();
            }
            catch (Exception)
            {
            }
        }
        private void cbxReverse_CheckedChanged(object sender, EventArgs e)
        {
            direction = !direction;
            //DoMouseClick();
            //Cursor.Position = new Point(Cursor.Position.X + 500, Cursor.Position.Y + 50);
            presentation.SlideShowSettings.Run();
        }


        private void frmClick(object sender, MouseEventArgs e)
        {
            
        }
    }
}
