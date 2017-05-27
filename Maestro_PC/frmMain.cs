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
        private PPT.Application pptApplication;
        private PPT.Presentation presentation;
        private PPT.Slide slide;
        private PPT.Slides slides;
        private int slidescount;

        Thread AcceptAndListeningThread;

        // helper variable
        Boolean isConnected = false;

        //bluetooth stuff
        BluetoothClient btClient;  //represent the bluetooth client connection
        BluetoothListener btListener; //represent this server bluetooth device

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
            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "Please check PowerPoint is running.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void slide_previous()
        {
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
                MetroFramework.MetroMessageBox.Show(this, "This page is the first page.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button_Previous_Click(object sender, EventArgs e)
        {
            slide_previous();
        }
        private void slide_next()
        {
            int slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                MetroFramework.MetroMessageBox.Show(this, "This page is the last page.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private void button_Next_Click(object sender, EventArgs e)
        {
            slide_next();
        }
        private void button_Mouse_Click(object sender, EventArgs e)
        {
            presentation.SlideShowWindow.View.LaserPointerEnabled = true;
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            this.Size = new Size(500, 500);
            MetroFramework.Controls.MetroButton button_Connect = new MetroFramework.Controls.MetroButton();
            button_Connect.Size = new Size(200, 50);
            button_Connect.Left = 150;
            button_Connect.Top = 120;
            button_Connect.Text = "Connect Powerpoint";
            button_Connect.Click += new EventHandler(button_Connect_Click);
            this.Controls.Add(button_Connect);


            MetroFramework.Controls.MetroButton button_Previous = new MetroFramework.Controls.MetroButton();
            button_Previous.Size = new Size(200, 50);
            button_Previous.Left = 150;
            button_Previous.Top = 200;
            button_Previous.Text = "Previous Slide";
            button_Previous.Click += new EventHandler(button_Previous_Click);
            this.Controls.Add(button_Previous);

            MetroFramework.Controls.MetroButton button_Next = new MetroFramework.Controls.MetroButton();
            button_Next.Size = new Size(200, 50);
            button_Next.Left = 150;
            button_Next.Top = 280;
            button_Next.Text = "Next Slide";
            button_Next.Click += new EventHandler(button_Next_Click);
            this.Controls.Add(button_Next);

            MetroFramework.Controls.MetroButton button_Mouse = new MetroFramework.Controls.MetroButton();
            button_Mouse.Size = new Size(200, 50);
            button_Mouse.Left = 150;
            button_Mouse.Top = 360;
            button_Mouse.Text = "Mouse Pointer";
            button_Mouse.Click += new EventHandler(button_Mouse_Click);
            this.Controls.Add(button_Mouse);

            //when the bluetooth is supported by this computer

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

                //creating and starting the thread
                AcceptAndListeningThread = new Thread(AcceptAndListen);

                AcceptAndListeningThread.Start();

            }
            else
            {
                UpdateLogText("Bluetooth not Supported!");
            }

        }
        StreamReader srReceiver;
        private void ReceiveMessages()
        {
            // Receive the response from the server
            srReceiver = new StreamReader(btClient.GetStream());

            // If the first character of the response is 1, connection was successful
            string ConResponse = srReceiver.ReadLine();

            // If the first character is a 1, connection was successful
            if (ConResponse[0] == '1')
            {
                // Update the form to tell it we are now connected
                this.Invoke(new UpdateLogCallback(this.UpdateLogText), new object[] { "Connected Successfully!" });
            }
            else // If the first character is not a 1 (probably a 0), the connection was unsuccessful
            {
                string Reason = "Not Connected: ";

                // Extract the reason out of the response message. The reason starts at the 3rd character
                Reason += ConResponse.Substring(2, ConResponse.Length - 2);

                // Exit the method
                return;
            }
            // While we are successfully connected, read incoming lines from the server
            while (isConnected)
            {
                // Show the messages in the log TextBox
                this.Invoke(new UpdateLogCallback(this.UpdateLogText), new object[] { srReceiver.ReadLine() });
            }
        }
        public void AcceptAndListen()
        {
            while (true)
            {
                if (isConnected)
                {
                    //TODO: if there is a device connected
                    //listening
                    try
                    {
                        UpdateLogTextFromThread("Listening….");
                        NetworkStream stream = btClient.GetStream();

                        Byte[] bytes = new Byte[512];

                        String retrievedMsg = "";

                        stream.Read(bytes, 0, 512);

                        stream.Flush();

                        for (int i = 0; i < bytes.Length; i++)
                        {
                            retrievedMsg += Convert.ToChar(bytes[i]);

                        }
                        //if(String.Equals(retrievedMsg, "1"))
                        //{
                        //    UpdateLogTextFromThread("1");
                        //    //slide_previous();
                        //}
                        //else if(String.Equals(retrievedMsg,"2"))
                        //{
                        //    UpdateLogTextFromThread("2");
                        //    //slide_next();
                        //}
                        UpdateLogTextFromThread(retrievedMsg);
                        if (!retrievedMsg.Contains("servercheck"))
                        {

                            //sendMessage("Message Received!");
                        }
                        //ReceiveMessages();
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




                    //TODO: if there is no connection
                    // accepting
                    //TODO: if there is no connection
                    // accepting
                    //TODO: if there is no connection
                    // accepting
                    try
                    {
                        btListener = new BluetoothListener(BluetoothService.RFCommProtocol);

                        UpdateLogTextFromThread("Listener created with TCP Protocol service " + BluetoothService.RFCommProtocol);
                        UpdateLogTextFromThread("Starting Listener….");
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
        public void UpdateLogText(String msg)
        {
            Console.WriteLine(msg);
            if(msg[msg.Length-1]== '1')
            {
                slide_previous();
            }
            else if(msg[msg.Length - 1] == '2')
            {
                slide_next();
            }
        }
        delegate void UpdateLogTextFromThreadDelegate(String msg);
        public void UpdateLogTextFromThread(String msg)
        {
            if (!this.IsDisposed && this.InvokeRequired)
            {
                this.Invoke(new UpdateLogTextFromThreadDelegate(UpdateLogText), new Object[] { msg });
            }
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                AcceptAndListeningThread.Abort();
                btClient.GetStream().Close();
                btClient.Dispose();
                btListener.Stop();
            }
            catch (Exception)
            {
            }
        }
    }
}
