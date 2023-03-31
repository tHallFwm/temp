using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;

namespace PowerVision.USBHid
{
    internal class HiddenForm : Form
    {
        #region Fields
        private Label label1;
        private USBInterface mUSBInterface = null;
        #endregion

        #region Methods
        public HiddenForm(USBInterface usbi)
        {
            mUSBInterface = usbi;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(-2000, -2000);
            this.Size = new System.Drawing.Size(1, 1);
            this.WindowState = FormWindowState.Minimized;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.ShowIcon = false;
            this.Load += new EventHandler(this.Load_Form);
            this.Activated += new EventHandler(this.Form_Activated);

        }

        private void Load_Form(object sender, EventArgs e)
        {
            InitializeComponent();

            this.Size = new System.Drawing.Size(5, 5);
            this.Hide();
        }

        private void Form_Activated(object sender, EventArgs e)
        {
            this.Visible = false;
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == DeviceManagement.WM_DEVICECHANGE && mUSBInterface != null)
                mUSBInterface.WndProc(ref m);
        }

        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(314, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "This is invisible form. To see USBInterface code click View Code";
            // 
            // HiddenForm
            // 
            this.ClientSize = new System.Drawing.Size(360, 80);
            this.Controls.Add(this.label1);
            this.Name = "HiddenForm";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
        #endregion
    }

    public class ConnectionStatusEventArgs : EventArgs
    {
        public bool ConnectionStatus { get; set; }
    }

    public class USBInterface : IDisposable
    {
        #region Fields
        private IntPtr deviceNotificationHandle;
        private SafeFileHandle hidHandle;
        private SafeFileHandle myDevicehidHandle;
        private string hidUsage;
        private bool myDeviceDetected;
        private bool multipleDevicesDetected;
        private string myDevicePathName;
        private int inputReportByteLength;
        private int outputReportByteLength;

        private DeviceManagement MyDeviceManagement = new DeviceManagement();
        private Hid MyHid = new Hid();

        private HiddenForm hiddenForm;
        #endregion

        #region Properties
        public event EventHandler<ConnectionStatusEventArgs> ConnectionStatusHandler;

        public bool ConnectionStatus { get { return myDeviceDetected; } }
        public bool ConnectionProblem { get { return multipleDevicesDetected; } }
        public int InputReportByteLength { get { return inputReportByteLength; } }
        public int OutputReportByteLength { get { return outputReportByteLength; } }
        public string ConnectedDevice { get; set; }
        #endregion

        #region Methods
 //       [CLSCompliant(false)]
        public USBInterface()
        {
            Guid hidGuid = Guid.Empty;
            bool success = false;

            ConnectedDevice = String.Empty;
            myDeviceDetected = false;
            FindTheHid();

            hiddenForm = new HiddenForm(this);
            hiddenForm.Show();

            Hid.HidD_GetHidGuid(ref hidGuid);
            success = MyDeviceManagement.RegisterForDeviceNotifications(string.Empty, hiddenForm.Handle, hidGuid, ref deviceNotificationHandle);
        }

        private bool FindTheHid()
        {
            bool deviceFound = false;
            Guid hidGuid = Guid.Empty;
            string[] devicePathName = new string[128];
            int memberIndex = 0;
            int numberOfInputBuffers = 0;
            bool success = false;

            myDeviceDetected = false;
            multipleDevicesDetected = false;
            Hid.HidD_GetHidGuid(ref hidGuid);
            deviceFound = MyDeviceManagement.FindDeviceFromGuid(hidGuid, ref devicePathName);
            if (deviceFound)
            {
                memberIndex = 0;

                // just spin thru all 80 of them.  Need to check for multiple devices
                while (memberIndex < devicePathName.Length)
                {
                    if (devicePathName[memberIndex] != null)
                    {
                        hidHandle = FileIO.CreateFile(devicePathName[memberIndex], 0, FileIO.FILE_SHARE_READ | FileIO.FILE_SHARE_WRITE, IntPtr.Zero, FileIO.OPEN_EXISTING, 0, 0);

                        if (!hidHandle.IsInvalid)
                        {
                            MyHid.DeviceAttributes.Size = Marshal.SizeOf(MyHid.DeviceAttributes);

                            success = Hid.HidD_GetAttributes(hidHandle, ref MyHid.DeviceAttributes);

                            if (success)
                            {
                                if ((MyHid.DeviceAttributes.VendorID == 0x084F && MyHid.DeviceAttributes.ProductID == 0x0001) ||    // EZView
                                    (MyHid.DeviceAttributes.VendorID == 0x2B28 && MyHid.DeviceAttributes.ProductID == 0x0001) ||    // Old C5 (MPC20)
                                    (MyHid.DeviceAttributes.VendorID == 0x2B28 && MyHid.DeviceAttributes.ProductID == 0x0003) ||    // CenturionNG - C5
                                    (MyHid.DeviceAttributes.VendorID == 0x2B28 && MyHid.DeviceAttributes.ProductID == 0x0005))      // EZView.r2
                                {
                                    // lock in the first valid one found
                                    if (myDeviceDetected == false)
                                    {
                                        myDeviceDetected = true;
                                        myDevicePathName = devicePathName[memberIndex];
                                        myDevicehidHandle = hidHandle;
                                    }
                                    else
                                    {
                                        multipleDevicesDetected = true;
                                    }
                                }
                                else
                                {
                                    hidHandle.Close();
                                }
                            }
                            else
                            {
                                hidHandle.Close();
                            }
                        }
                    }
                    memberIndex++;
                }
            }
            //restore handle to my device
            hidHandle = myDevicehidHandle;

            if (myDeviceDetected)
            {
                MyHid.Capabilities = MyHid.GetDeviceCapabilities(hidHandle);

                inputReportByteLength = MyHid.Capabilities.InputReportByteLength;
                outputReportByteLength = MyHid.Capabilities.OutputReportByteLength;

                hidUsage = MyHid.GetHidUsage(MyHid.Capabilities);
                success = MyHid.GetNumberOfInputBuffers(hidHandle, ref numberOfInputBuffers);
            }
            SendConnectionStatusEvent(myDeviceDetected);

            return myDeviceDetected;
        }

        public bool ReadInputReport(ref byte[] inputReportBuffer)
        {
            bool success = false;

            if (!hidHandle.IsInvalid)
            {
                if (MyHid.Capabilities.InputReportByteLength > 0)
                {
                    if (hidHandle == null || hidHandle.IsClosed || hidHandle.IsInvalid)
                        throw new USBCommunicationException();
                    success = Hid.HidD_GetInputReport(hidHandle, inputReportBuffer, inputReportBuffer.Length);
                    if (!success)
                        throw new USBCommunicationException();
                }
            }

            return success;
        }

        public bool WriteOutputReport(byte[] outputReportData)
        {
            bool success = false;

            if (!hidHandle.IsInvalid)
            {
                if (MyHid.Capabilities.OutputReportByteLength > 0)
                {
                    if (hidHandle == null || hidHandle.IsClosed || hidHandle.IsInvalid)
                        throw new USBCommunicationException();
                    success = Hid.HidD_SetOutputReport(hidHandle, outputReportData, outputReportData.Length + 1);
                    if (!success)
                        throw new USBCommunicationException();
                }
            }

            return success;
        }

        private void CloseCommunications()
        {
            if (hidHandle != null && !hidHandle.IsInvalid)
                hidHandle.Close();

            myDeviceDetected = false;
            SendConnectionStatusEvent(myDeviceDetected);
        }

        public void Reset()
        {
            CloseCommunications();
            FindTheHid();
        }

        public void Shutdown()
        {
            CloseCommunications();
            MyDeviceManagement.StopReceivingDeviceNotifications(deviceNotificationHandle);
        }

        internal void WndProc(ref Message m)
        {
            if (m.Msg == DeviceManagement.WM_DEVICECHANGE)
            {
                if (m.WParam.ToInt32() == DeviceManagement.DBT_DEVICEARRIVAL)
                {
                    //if (!myDeviceDetected)
                    {
                        FindTheHid();
                    }
                }
                else if (m.WParam.ToInt32() == DeviceManagement.DBT_DEVICEREMOVECOMPLETE)
                {
                    if (MyDeviceManagement.DeviceNameMatch(m, myDevicePathName) && this.myDeviceDetected)
                    {
                        CloseCommunications();
                    }
                }
            }
        }

        public void Dispose()
        {

        }

        private void SendConnectionStatusEvent(bool status)
        {
            if (ConnectionStatusHandler != null)
            {
                ConnectionStatusEventArgs args = new ConnectionStatusEventArgs();
                args.ConnectionStatus = status;
                ConnectionStatusHandler(this, args);
            }
        }
        #endregion
    }

    public class USBCommunicationException : System.Exception
    {
        public USBCommunicationException() : base() { }
        public USBCommunicationException(string message) : base(message) { }
        public USBCommunicationException(string message, System.Exception inner) : base(message, inner) { }
    }

}
