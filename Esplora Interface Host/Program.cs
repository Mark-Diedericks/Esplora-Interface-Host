using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Esplora_Interface_Host
{
    public class Program
    {
        private static Program prog;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            prog = new Program();
            Application.Run();
        }

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private Timer timer;
        private SerialPort serial;
        private SimpleHTTPServer server;
        private static List<string> commands;

        public Program()
        {
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);
            trayIcon = new NotifyIcon();
            trayIcon.Icon = Properties.Resources.arduino_icon_2;
            trayIcon.Text = "Esplora Interface Host";
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;

            commands = new List<string>();

            commands.Add("crossdomain.xml");
            commands.Add("poll");
            commands.Add("A");
            commands.Add("E");
            commands.Add("C");
            commands.Add("D");
            commands.Add("I");
            commands.Add("J");
            commands.Add("K");
            commands.Add("X");
            commands.Add("Y");
            commands.Add("Z");
            commands.Add("R");
            commands.Add("G");
            commands.Add("B");
            commands.Add("S");
            commands.Add("L");
            commands.Add("M");
            commands.Add("T");
            commands.Add("N");

            server = new SimpleHTTPServer(63287);

            serial = new SerialPort();
            serial.BaudRate = 57600;
            serial.PortName = AutodetectArduinoPort();
            serial.DtrEnable = true;
            serial.RtsEnable = true;
            serial.DataReceived += new SerialDataReceivedEventHandler(Serial_DataReceived);
            serial.Open();

            timer = new Timer();
            timer.Interval = 10;
            timer.Tick += new EventHandler(Timer_Tick);
            timer.Enabled = true;
            timer.Start();
        }

        public static Program instance()
        {
            return prog;
        }

        public static List<string> getCommands()
        {
            return commands;
        }

        private string AutodetectArduinoPort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains("Arduino"))
                    {
                        return deviceId;
                    }
                }
            }
            catch (ManagementException e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
            }


            if(MessageBox.Show("Could not find an arduino on any USB ports.", "Error") != DialogResult.None)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
                trayMenu.Dispose();
                Application.Exit();
            }

            return null;
        }

        public int A = 0;
        public int E = 0;
        public int C = 0;
        public int D = 0;

        public int I = 0;
        public int J = 0;
        public int K = 0;

        public int X = 0;
        public int Y = 0;
        public int Z = 0;

        public int R = 255;
        public int G = 255;
        public int B = 255;

        public int S = 0;

        public int L = 0;

        public int M = 0;

        public int T = 0;

        public int N = 0;

        private void Serial_DataReceived(object sender, SerialDataReceivedEventArgs ea)
        {
            if (!serial.IsOpen) serial.Open();
            if (!serial.IsOpen) return;

            string line = serial.ReadLine();

            switch (line.ToCharArray()[0])
            {
                case 'a':
                case 'A':
                    //Switch 1
                    try { A = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'e':
                case 'E':
                    //Switch 2
                    try { E = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'c':
                case 'C':
                    //Switch 3
                    try { C = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'd':
                case 'D':
                    //Switch 4
                    try { D = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'i':
                case 'I':
                    //Joystick Button
                    try { I = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'j':
                case 'J':
                    //Joystick X
                    try { J = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'k':
                case 'K':
                    //Joystick Y
                    try { K = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'x':
                case 'X':
                    //Accelerometer X
                    try { X = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'y':
                case 'Y':
                    //Accelerometer Y
                    try { Y = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'z':
                case 'Z':
                    //Accelerometer Z
                    try { Z = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 's':
                case 'S':
                    //Linear Potentiometer (Slider)
                    try { S = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'l':
                case 'L':
                    //Light Sensor
                    try { L = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 'm':
                case 'M':
                    //Microphone
                    try { M = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
                case 't':
                case 'T':
                    //Temperature Sensor
                    try { T = Convert.ToInt32(Regex.Replace(line, "[^0-9.+-]", "")); }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                    break;
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            string id = "";

            if (!serial.IsOpen) serial.Open();

            if (serial.IsOpen)
            {
                //Buttons
                try
                { 
                    id = "A";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "E";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }
                
                try
                {
                    id = "C";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "D";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //Joystick
                try
                {
                    id = "I";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "J";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "K";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //Accelerometer
                try
                {
                    id = "X";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "Y";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "Z";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //RGB LED
                try
                {
                    id = "R";
                    serial.WriteLine(id.ToUpper() + R.ToString());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "G";
                    serial.WriteLine(id.ToUpper() + G.ToString());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                try
                {
                    id = "B";
                    serial.WriteLine(id.ToUpper() + B.ToString());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //Slider
                try
                {
                    id = "S";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //Light Sensor
                try
                {
                    id = "L";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //Microphone
                try
                {
                    id = "M";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //Temperature Sensor
                try
                {
                    id = "T";
                    serial.WriteLine(id.ToUpper());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }

                //Buzzer
                try
                {
                    id = "N";
                    serial.WriteLine(id.ToUpper() + N.ToString());
                } catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(id.ToUpper() + ": " + ex.Message);
                }
            }
        }

        public string poll()
        {
            string result = "";
            
            result += "A " + A.ToString() + System.Environment.NewLine;
            result += "E " + E.ToString() + System.Environment.NewLine;
            result += "C " + C.ToString() + System.Environment.NewLine;
            result += "D " + D.ToString() + System.Environment.NewLine;

            result += "I " + I.ToString() + System.Environment.NewLine;
            result += "J " + J.ToString() + System.Environment.NewLine;
            result += "K " + K.ToString() + System.Environment.NewLine;

            result += "X " + X.ToString() + System.Environment.NewLine;
            result += "Y " + Y.ToString() + System.Environment.NewLine;
            result += "Z " + Z.ToString() + System.Environment.NewLine;

            result += "S " + S.ToString() + System.Environment.NewLine;

            result += "L " + L.ToString() + System.Environment.NewLine;

            result += "M " + M.ToString() + System.Environment.NewLine;

            result += "T " + T.ToString();

            return result;
        }

        private void OnExit(object sender, EventArgs e)
        {
            if (serial.IsOpen) serial.Close();
            trayIcon.Visible = false;
            trayIcon.Dispose();
            trayMenu.Dispose();
            Environment.Exit(0);
            Application.Exit();
        }
    }
}
