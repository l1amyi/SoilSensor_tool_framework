using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
namespace Test
{
    public partial class Form1 : Form
    {
        private SerialPort serialPort;
        private const byte FRAME_HEAD_1 = 0x7D;
        private const byte FRAME_HEAD_2 = 0x95;
        private const byte FRAME_TAIL_1 = 0x6A;
        private const byte FRAME_TAIL_2 = 0xE2;
        private const uint ADDR_DEVICE = 0x04004200;
        private const uint ADDR_BOARDCAST = 0x7F7F7F7F;
        private const uint ADDR_UPPER = 0x1F162C6A;
        private const uint ADDR_ADMIN = 0xA62E59D7;
        private const byte CMD_RESET_DEVICE = 0x04;
        private const byte CMD_QUERY_DEVICE_ADDR = 0x10;
        private const byte CMD_QUERY_DEVICE_ADDR_REPLY = 0x80;
        private const byte CMD_QUERY_DEVICE_INFO = 0x11;
        private const byte CMD_QUERY_DEVICE_INFO_REPLY = 0x81;
        private const byte CMD_QUERY_SENSOR_MODEL = 0x12;
        private const byte CMD_QUERY_SENSOR_MODEL_REPLY = 0x82;
        private const byte CMD_QUERY_RTC = 0x13;
        private const byte CMD_QUERY_RTC_REPLY = 0x83;
        private const byte CMD_QUERY_TIME_INTERVAL = 0x14;
        private const byte CMD_QUERY_TIME_INTERVAL_REPLY = 0x84;
        private const byte CMD_QUERY_REAL_TIME_SENSOR_DATA = 0x15;
        private const byte CMD_QUERY_REAL_TIME_SENSOR_DATA_REPLY = 0x85;
        private const byte CMD_QUERY_COLLECTION_START_TIME = 0x16;
        private const byte CMD_QUERY_COLLECTION_START_TIME_REPLY = 0x86;
        private const byte CMD_QUERY_COLLECTION_DATA = 0x17;
        private const byte CMD_QUERY_COLLECTION_DATA_REPLY = 0x87;
        private const byte CMD_SET_DEVICE_SETTING = 0x21;
        private const byte CMD_SET_DEVICE_SETTING_REPLY = 0x91;
        private const byte CMD_SET_SENSOR_MODEL = 0x22;
        private const byte CMD_SET_SENSOR_MODEL_REPLY = 0x92;
        private const byte CMD_SET_RTC = 0x23;
        private const byte CMD_SET_RTC_REPLY = 0x93;
        private const byte CMD_SET_TIME_INTERVAL = 0x24;
        private const byte CMD_SET_TIME_INTERVAL_REPLY = 0x94;
        private const byte CMD_COLLECTION_RESTART = 0x25;
        private const byte CMD_COLLECTION_RESTART_REPLY = 0x95;
        private const byte CMD_CLEAR_COLLECTION_DATA = 0x26;
        private const byte CMD_CLEAR_COLLECTION_DATA_REPLY = 0x96;
        private const byte CMD_UPGRADE_DEVICE = 0x50;
        private const byte CMD_UPGRADE_DEVICE_REPLY = 0xC0;

        private const uint PASSWORD = 0x23F209C3;

        private Timer timerCurrentTime;

        public Form1()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // 固定边框

            serialPort = new SerialPort();
            serialPort.BaudRate = 115200;
            //serialPort.DataReceived += SerialPort_DataReceived;
            this.Load += Form1_Load; // 注册窗体加载事件
            this.Text = "SoilSensorTool_V1.1.0";

            tabControl1.Alignment = TabAlignment.Left;
            tabControl1.DrawMode = TabDrawMode.OwnerDrawFixed;
            tabControl1.ItemSize = new Size(50, 100);
            tabControl1.DrawItem += (s, e) =>
            {
                TabPage page = tabControl1.TabPages[e.Index];
                Rectangle rect = e.Bounds;
                StringFormat sf = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(page.Text, tabControl1.Font, Brushes.Black, rect, sf);
            };

            tabControl1.DrawItem += (s, e) =>
            {
                TabControl tab = (TabControl)s;
                for (int i = 0; i < tab.TabPages.Count; i++)
                {
                    Rectangle rect = tab.GetTabRect(i);
                    bool selected = (i == tab.SelectedIndex);

                    // 选中和未选中的背景色
                    Color backColor = selected ? Color.LightSkyBlue : Color.LightGray;
                    Color foreColor = selected ? Color.Black : Color.Gray;

                    using (SolidBrush brush = new SolidBrush(backColor))
                    {
                        e.Graphics.FillRectangle(brush, rect);
                    }

                    // 文字居中
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;

                    using (SolidBrush brush = new SolidBrush(foreColor))
                    {
                        e.Graphics.DrawString(tab.TabPages[i].Text, tab.Font, brush, rect, sf);
                    }
                }
                e.DrawFocusRectangle();
            };

            // 定时器
            timerCurrentTime = new Timer();
            timerCurrentTime.Interval = 1000; // 1秒刷新一次
            timerCurrentTime.Tick += TimerCurrentTime_Tick;
            timerCurrentTime.Start();

            //传感器序号
            for (int i = 0; i < 10; i++)
            {
                comboboxSensorIdx.Items.Add($"{i + 1}#传感器");
            }

            comboboxSensorIdx.SelectedIndex = 0;

            dataGridViewSensor.ReadOnly = false;
            dataGridViewSensor.AllowUserToAddRows = false;


            for (int i = 0; i < 10; i++)
            {
                ComboBox comboBox = this.Controls.Find("comboBoxSensorModel" + (i + 1), true).FirstOrDefault() as ComboBox;

                if (comboBox != null)
                {
                    comboBox.Items.Add("未配置");
                    comboBox.Items.Add("PL30传感器");
                }
            }



        }

        private void ScanSerialPorts()
        {
            string[] ports = SerialPort.GetPortNames();
            com.Items.Clear();
            com.Items.AddRange(ports);
            if (com.Items.Count > 0)
                com.SelectedIndex = 0;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ScanSerialPorts();
            dataGridViewSensor.RowPostPaint += dataGridViewSensor_RowPostPaint;
        }

        private void dataGridViewSensor_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            using (SolidBrush brush = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                string rowIndex = (e.RowIndex + 1).ToString();
                e.Graphics.DrawString(rowIndex, dgv.Font, brush,
                    e.RowBounds.Left + 10, e.RowBounds.Top + 4);
            }
        }

        private void OpenPort()
        {
            if (!serialPort.IsOpen)
            {
                serialPort.Open();
            }
        }

        private void ClosePort()
        {
            if (serialPort.IsOpen)
            {
                serialPort.Close();
            }
        }

        private void SendHex(byte[] bytes)
        {
            if (serialPort.IsOpen)
            {
                serialPort.DiscardInBuffer(); // 清空输入缓冲区
                serialPort.Write(bytes, 0, bytes.Length);

                string log;
                log = "发送数据: ";

                for (int i = 0; i < bytes.Length; i++)
                {
                    log += bytes[i].ToString("X2") + " ";
                }

                AddLogToTextBox(log);
            }
        }



        private byte CalculateChecksum(byte[] buffer, int size)
        {
            byte checksum = 0;
            for (int i = 0; i < size; i++)
            {
                checksum += buffer[i];
            }
            checksum = (byte)~checksum;
            return checksum;
        }

        private void connect_button_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                serialPort.PortName = com.SelectedItem.ToString();
                byte[] query_address = LoadMsg(
                CMD_QUERY_DEVICE_ADDR,
                ADDR_BOARDCAST,
                ADDR_UPPER,
                null,
                0);

                PrintHex(query_address);
                OpenPort();
                SendHex(query_address);

                byte cmd = 0;
                uint targetAddress = 0;
                uint sourceAddress = 0;
                byte[] content;

                ReceiveMsg(out targetAddress, out sourceAddress, out cmd, out content);

                if (cmd == CMD_QUERY_DEVICE_ADDR_REPLY)
                {
                    // 处理查询设备地址的响应

                    if (content != null && content.Length >= 4)
                    {
                        uint deviceAddress = (uint)((content[0] << 24) |
                                                      (content[1] << 16) |
                                                      (content[2] << 8) |
                                                      content[3]);
                        Console.WriteLine($"设备地址: {deviceAddress:X8}");

                        for (int i = 0; i < comboBoxDevAdd.Items.Count; i++)
                        {
                            if (comboBoxDevAdd.Items[i].ToString() == deviceAddress.ToString("X8"))
                            {
                                comboBoxDevAdd.SelectedIndex = i;
                                connect_button.Text = "断开端口";
                                return; // 如果设备地址已存在，则直接返回
                            }
                        }

                        comboBoxDevAdd.Items.Add(deviceAddress.ToString("X8"));

                        if (comboBoxDevAdd.Items.Count == 1)
                        {
                            comboBoxDevAdd.SelectedIndex = 0;
                            connect_button.Text = "断开端口";
                        }
                        else
                        {
                            comboBoxDevAdd.SelectedIndex = comboBoxDevAdd.Items.Count - 1;
                            connect_button.Text = "断开端口";
                        }
                    }
                    else
                    {
                        Console.WriteLine("响应内容格式错误");
                    }
                }
            }
            else
            {
                ClosePort();
                connect_button.Text = "连接端口";
            }



        }

        private bool ReceiveMsg(out uint targetAddress, out uint sourceAddress, out byte cmd, out byte[] content, int timeout = 1000)
        {
            try
            {
                serialPort.ReadTimeout = timeout; // 设置读取超时为1秒


                byte MsgHead1 = (byte)serialPort.ReadByte();
                byte MsgHead2 = (byte)serialPort.ReadByte();


                if (MsgHead1 != FRAME_HEAD_1 || MsgHead2 != FRAME_HEAD_2)
                {
                    Console.WriteLine("帧头错误");
                    targetAddress = 0;
                    sourceAddress = 0;
                    cmd = 0;
                    content = null;

                    string log;
                    log = "帧头错误: " + MsgHead1.ToString("X2") + " " + MsgHead2.ToString("X2");

                    AddLogToTextBox(log);

                    return false;
                }

                cmd = (byte)serialPort.ReadByte();

                byte len = (byte)serialPort.ReadByte();

                targetAddress = (uint)((serialPort.ReadByte() << 24) |
                                       (serialPort.ReadByte() << 16) |
                                       (serialPort.ReadByte() << 8) |
                                       serialPort.ReadByte());

                sourceAddress = (uint)((serialPort.ReadByte() << 24) |
                                       (serialPort.ReadByte() << 16) |
                                       (serialPort.ReadByte() << 8) |
                                       serialPort.ReadByte());

                Console.WriteLine($"收到命令: {cmd:X2}, 目标地址: {targetAddress:X8}, 源地址: {sourceAddress:X8}, 内容长度: {len}");

                int rx_len = 0;
                if (len > 0)
                {
                    content = new byte[len];
                    while (rx_len < len)
                    {
                        rx_len += serialPort.Read(content, rx_len, len - rx_len);
                    }
                }
                else
                {
                    content = null;
                }

                if (content != null)
                {
                    if (rx_len != content.Length)
                    {
                        Console.WriteLine("内容长度错误 " + rx_len + " " + len);
                        targetAddress = 0;
                        sourceAddress = 0;
                        cmd = 0;
                        content = null;

                        string log;
                        log = "内容长度错误: " + rx_len + " " + len;
                        AddLogToTextBox(log);

                        return false;
                    }
                }


                byte checksum = (byte)serialPort.ReadByte();

                byte MsgTail1 = (byte)serialPort.ReadByte();

                byte MsgTail2 = (byte)serialPort.ReadByte();

                if (MsgTail1 != FRAME_TAIL_1 || MsgTail2 != FRAME_TAIL_2)
                {
                    Console.WriteLine("帧尾错误");
                    targetAddress = 0;
                    sourceAddress = 0;
                    cmd = 0;
                    content = null;

                    string log;
                    log = "帧尾错误: " + MsgTail1.ToString("X2") + " " + MsgTail2.ToString("X2");
                    AddLogToTextBox(log);

                    return false;
                }

                byte[] rx_msg = new byte[12 + len + 3];

                rx_msg[0] = MsgHead1;
                rx_msg[1] = MsgHead2;
                rx_msg[2] = cmd;
                rx_msg[3] = len;
                rx_msg[4] = (byte)((targetAddress >> 24) & 0xFF);
                rx_msg[5] = (byte)((targetAddress >> 16) & 0xFF);
                rx_msg[6] = (byte)((targetAddress >> 8) & 0xFF);
                rx_msg[7] = (byte)(targetAddress & 0xFF);
                rx_msg[8] = (byte)((sourceAddress >> 24) & 0xFF);
                rx_msg[9] = (byte)((sourceAddress >> 16) & 0xFF);
                rx_msg[10] = (byte)((sourceAddress >> 8) & 0xFF);
                rx_msg[11] = (byte)(sourceAddress & 0xFF);
                if (len > 0 && content != null)
                {
                    Array.Copy(content, 0, rx_msg, 12, len);
                }
                rx_msg[12 + len] = checksum;
                rx_msg[12 + len + 1] = MsgTail1;
                rx_msg[12 + len + 2] = MsgTail2;

                string log_msg = "收到消息:";
                for (int i = 0; i < rx_msg.Length; i++)
                {
                    log_msg += " " + rx_msg[i].ToString("X2");
                }
                AddLogToTextBox(log_msg);

                if (checksum != CalculateChecksum(rx_msg, 12 + len))
                {
                    Console.WriteLine($"收到校验和: {checksum:X2}");
                    Console.WriteLine("校验和错误");
                    targetAddress = 0;
                    sourceAddress = 0;
                    cmd = 0;
                    content = null;

                    string log;
                    log = " 校验和错误: " + checksum.ToString("X2") + " " + CalculateChecksum(rx_msg, 12 + len).ToString("X2");
                    AddLogToTextBox(log);

                    return false;
                }

                return true;
            }
            catch (TimeoutException)
            {
                Console.WriteLine("接收超时");
                targetAddress = 0;
                sourceAddress = 0;
                cmd = 0;
                content = null;

                string log;
                log = "接收超时";
                AddLogToTextBox(log);

                return false;
            }

            return true;
        }

        // 组帧
        private byte[] LoadMsg(
            byte cmd,
            uint targetAddress,
            uint sourceAddr,
            byte[] content,
            byte len)
        {
            byte[] msg = new byte[12 + len + 3];

            msg[0] = FRAME_HEAD_1;
            msg[1] = FRAME_HEAD_2;

            msg[2] = cmd;
            msg[3] = len;

            msg[4] = (byte)((targetAddress >> 24) & 0xFF);
            msg[5] = (byte)((targetAddress >> 16) & 0xFF);
            msg[6] = (byte)((targetAddress >> 8) & 0xFF);
            msg[7] = (byte)(targetAddress & 0xFF);

            msg[8] = (byte)((sourceAddr >> 24) & 0xFF);
            msg[9] = (byte)((sourceAddr >> 16) & 0xFF);
            msg[10] = (byte)((sourceAddr >> 8) & 0xFF);
            msg[11] = (byte)(sourceAddr & 0xFF);

            if (len > 0 && content != null)
            {
                Array.Copy(content, 0, msg, 12, len);
            }

            msg[12 + len] = CalculateChecksum(msg, 12 + len);

            msg[12 + len + 1] = FRAME_TAIL_1;
            msg[12 + len + 2] = FRAME_TAIL_2;

            return msg;
        }

        // 解帧
        private bool UnloadMsg(
            byte[] msg,
            out byte cmd,
            out uint targetAddress,
            out uint sourceAddr,
            byte[] content,
            out byte len)
        {
            cmd = 0;
            targetAddress = 0;
            sourceAddr = 0;
            len = 0;

            if (msg[0] != FRAME_HEAD_1 || msg[1] != FRAME_HEAD_2)
                return false;

            len = msg[3];
            if (msg[12 + len + 1] != FRAME_TAIL_1 || msg[12 + len + 2] != FRAME_TAIL_2)
                return false;

            cmd = msg[2];

            targetAddress = (uint)((msg[4] << 24) | (msg[5] << 16) | (msg[6] << 8) | msg[7]);
            sourceAddr = (uint)((msg[8] << 24) | (msg[9] << 16) | (msg[10] << 8) | msg[11]);

            if (len > 0 && content != null)
            {
                Array.Copy(msg, 12, content, 0, len);
            }

            return true;
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            int bytesToRead = serialPort.BytesToRead;
            byte[] buffer = new byte[bytesToRead];
            serialPort.Read(buffer, 0, bytesToRead);

            // 转换为十六进制字符串
            string hex = BitConverter.ToString(buffer).Replace("-", " ");

            // 跨线程更新UI
            this.Invoke(new Action(() =>
            {
                MessageBox.Show("收到数据: " + hex);
            }));
        }

        private void PrintHex(byte[] data)
        {
            if (data == null || data.Length == 0)
                return;

            StringBuilder sb = new StringBuilder();

            Console.Write("len " + data.Length + " bytes: ");

            foreach (byte b in data)
            {
                sb.Append(b.ToString("X2") + " ");
            }
            Console.WriteLine(sb.ToString().Trim());
        }

        private void btQueryDevInfo_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            Console.WriteLine(comboBoxDevAdd.Text);

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddress = ADDR_UPPER;
            byte[] queryDeviceInfo = LoadMsg(
                CMD_QUERY_DEVICE_INFO,
                targetAddress,
                sourceAddress,
                null,
                0);
            PrintHex(queryDeviceInfo);
            SendHex(queryDeviceInfo);
            byte cmd = 0;
            uint receivedTargetAddress = 0;
            uint receivedSourceAddress = 0;
            byte[] content;
            if (ReceiveMsg(out receivedTargetAddress, out receivedSourceAddress, out cmd, out content))
            {
                if (cmd == CMD_QUERY_DEVICE_INFO_REPLY)
                {
                    // 处理查询设备信息的响应
                    if (content != null && content.Length >= 4)
                    {
                        PrintHex(content);

                        uint sn = (uint)((content[0] << 24) | (content[1] << 16) | (content[2] << 8) | content[3]);
                        uint ver = (uint)((content[4] << 24) | (content[5] << 16) | (content[6] << 8) | content[7]);
                        byte sensor_num = content[8];
                        byte battery_level = content[9];

                        SafeSetText(textSerialNumber, sn.ToString());
                        SafeSetText(textVersion, ver.ToString("X8"));
                        SafeSetText(textSensorNum, sensor_num.ToString());
                        SafeSetText(textBetteryLevel, battery_level.ToString());

                    }
                    else
                    {
                        MessageBox.Show("响应内容格式错误");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void btSetDevInfo_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            uint sn = textSerialNumber.Text.Length > 0 ? Convert.ToUInt32(textSerialNumber.Text) : 0;
            byte sensor_num = textSensorNum.Text.Length > 0 ? Convert.ToByte(textSensorNum.Text) : (byte)0;

            Console.WriteLine($"设置设备信息: SN={sn}, SensorNum={sensor_num}");

            byte[] content = new byte[5];

            content[0] = (byte)((sn >> 24) & 0xFF);
            content[1] = (byte)((sn >> 16) & 0xFF);
            content[2] = (byte)((sn >> 8) & 0xFF);
            content[3] = (byte)(sn & 0xFF);
            content[4] = sensor_num;

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddress = ADDR_UPPER;

            byte[] setDeviceInfo = LoadMsg(
                CMD_SET_DEVICE_SETTING,
                targetAddress,
                sourceAddress,
                content,
                (byte)content.Length);
            PrintHex(setDeviceInfo);
            SendHex(setDeviceInfo);

            byte cmd = 0;
            uint receivedTargetAddress = 0;
            uint receivedSourceAddress = 0;
            byte[] responseContent;

            if (ReceiveMsg(out receivedTargetAddress, out receivedSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_SET_DEVICE_SETTING_REPLY)
                {
                    PrintHex(responseContent);

                    // 处理设置设备信息的响应
                    if (responseContent.Length == 11 && responseContent[0] == 1)
                    {
                        MessageBox.Show("设备信息设置成功");
                    }
                    else
                    {
                        MessageBox.Show("响应内容格式错误");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void btRequestRtc_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddress = ADDR_UPPER;

            byte[] setDeviceInfo = LoadMsg(
                CMD_QUERY_RTC,
                targetAddress,
                sourceAddress,
                null,
                0);
            PrintHex(setDeviceInfo);
            SendHex(setDeviceInfo);

            byte cmd = 0;
            uint receivedTargetAddress = 0;
            uint receivedSourceAddress = 0;
            byte[] responseContent;

            if (ReceiveMsg(out receivedTargetAddress, out receivedSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_QUERY_RTC_REPLY)
                {
                    PrintHex(responseContent);

                    if (responseContent.Length == 4)
                    {
                        uint utc_seconds = responseContent[0];
                        utc_seconds = (utc_seconds << 8) | responseContent[1];
                        utc_seconds = (utc_seconds << 8) | responseContent[2];
                        utc_seconds = (utc_seconds << 8) | responseContent[3];

                        DateTime utcTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(utc_seconds);

                        SafeSetText(lbRtc, utc_seconds + " " + utcTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                    else
                    {
                        MessageBox.Show("响应内容格式错误");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void btSetRtc_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            byte[] content = new byte[4];
            DateTime now = DateTime.Now;
            uint utc_seconds = (uint)(now - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;

            content[0] = (byte)((utc_seconds >> 24) & 0xFF);
            content[1] = (byte)((utc_seconds >> 16) & 0xFF);
            content[2] = (byte)((utc_seconds >> 8) & 0xFF);
            content[3] = (byte)(utc_seconds & 0xFF);

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;

            byte[] setRtcMsg = LoadMsg(
                CMD_SET_RTC,
                targetAddress,
                sourceAddr,
                content,
                (byte)content.Length);

            PrintHex(setRtcMsg);
            SendHex(setRtcMsg);

            byte cmd = 0;
            uint responseTargetAddress = 0;
            uint responseSourceAddress = 0;

            byte[] responseContent;

            if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_SET_RTC_REPLY)
                {
                    PrintHex(responseContent);

                    // 处理设置RTC的响应
                    if (responseContent.Length == 5 && responseContent[0] == 1)
                    {
                        MessageBox.Show("RTC设置成功");
                    }
                    else
                    {
                        MessageBox.Show("响应内容格式错误");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }

        }

        private void TimerCurrentTime_Tick(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            uint utc_seconds = (uint)(now - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
            lbCurrentTime.Text = utc_seconds + " " + now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private void btRequestTimeIntval_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddress = ADDR_UPPER;

            byte[] setDeviceInfo = LoadMsg(
                CMD_QUERY_TIME_INTERVAL,
                targetAddress,
                sourceAddress,
                null,
                0);
            PrintHex(setDeviceInfo);
            SendHex(setDeviceInfo);

            byte cmd = 0;
            uint receivedTargetAddress = 0;
            uint receivedSourceAddress = 0;
            byte[] responseContent;

            if (ReceiveMsg(out receivedTargetAddress, out receivedSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_QUERY_TIME_INTERVAL_REPLY)
                {
                    PrintHex(responseContent);

                    if (responseContent.Length == 2)
                    {
                        uint timeInterval = (uint)(responseContent[0] << 8) | responseContent[1];

                        SafeSetText(textTimeIntval, timeInterval.ToString());
                    }
                    else
                    {
                        MessageBox.Show("响应内容格式错误");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void btSetTimeIntval_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            ushort timeInterval = textTimeIntval.Text.Length > 0 ? Convert.ToUInt16(textTimeIntval.Text) : (ushort)0;

            byte[] content = new byte[2];

            content[0] = (byte)((timeInterval >> 8) & 0xFF);
            content[1] = (byte)(timeInterval & 0xFF);

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;

            byte[] setRtcMsg = LoadMsg(
                CMD_SET_TIME_INTERVAL,
                targetAddress,
                sourceAddr,
                content,
                (byte)content.Length);

            PrintHex(setRtcMsg);
            SendHex(setRtcMsg);

            byte cmd = 0;
            uint responseTargetAddress = 0;
            uint responseSourceAddress = 0;

            byte[] responseContent;

            if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_SET_TIME_INTERVAL_REPLY)
                {
                    PrintHex(responseContent);

                    // 处理设置时间间隔的响应
                    if (responseContent.Length == 3 && responseContent[0] == 1)
                    {
                        MessageBox.Show("时间间隔设置成功");
                    }
                    else
                    {
                        MessageBox.Show("响应内容格式错误");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void SafeSetText(Control ctrl, string text)
        {
            if (ctrl.InvokeRequired)
            {
                ctrl.Invoke(new Action(() => ctrl.Text = text));
            }
            else
            {
                ctrl.Text = text;
            }
        }

        private void btQueryStartTimeCount_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            if (comboboxSensorIdx.SelectedIndex < 0)
            {
                MessageBox.Show("请选择传感器序号");
                return;
            }

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;

            byte[] content = new byte[1];
            content[0] = (byte)comboboxSensorIdx.SelectedIndex;


            byte[] queryMsg = LoadMsg(
                CMD_QUERY_COLLECTION_START_TIME,
                targetAddress,
                sourceAddr,
                content,
                (byte)content.Length);

            PrintHex(queryMsg);
            SendHex(queryMsg);

            byte cmd = 0;
            uint responseTargetAddress = 0;
            uint responseSourceAddress = 0;

            byte[] responseContent;

            if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_QUERY_COLLECTION_START_TIME_REPLY)
                {
                    PrintHex(responseContent);

                    // 处理查询启动时间计数的响应
                    if (responseContent.Length == 8)
                    {
                        uint startTimeCount = (uint)(responseContent[0] << 24 | responseContent[1] << 16 | responseContent[2] << 8 | responseContent[3]);
                        uint count = (uint)(responseContent[4] << 24 | responseContent[5] << 16 | responseContent[6] << 8 | responseContent[7]);
                        string dataTime = DateTimeOffset.FromUnixTimeSeconds(startTimeCount).ToString("yyyy-MM-dd HH:mm:ss");
                        SafeSetText(textStartTime, dataTime);
                        SafeSetText(textRecordCount, count.ToString());
                    }
                    else
                    {
                        MessageBox.Show("响应内容格式错误");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }


        private void btExportData_Click(object sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "请选择导出文件名";
                dialog.Filter = "Excel文件|*.xlsx";
                dialog.FileName = "土壤数据" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ExportDataGridViewToXlsx(dataGridViewSensor, dialog.FileName);
                    MessageBox.Show("导出成功！\n" + dialog.FileName);
                }
            }
        }

        private void ExportDataGridViewToXlsx(DataGridView dgv, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // 写入表头
                for (int col = 0; col < dgv.Columns.Count; col++)
                {
                    worksheet.Cell(1, col + 1).Value = dgv.Columns[col].HeaderText;
                }

                // 写入内容
                for (int row = 0; row < dgv.Rows.Count; row++)
                {
                    if (dgv.Rows[row].IsNewRow) continue;
                    for (int col = 0; col < dgv.Columns.Count; col++)
                    {
                        var cellValue = dgv.Rows[row].Cells[col].Value;
                        worksheet.Cell(row + 2, col + 1).Value = cellValue == null ? "" : cellValue.ToString();
                    }
                }

                workbook.SaveAs(filePath);
            }
        }

        private void btQuerySensorData_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            if (comboboxSensorIdx.SelectedIndex < 0)
            {
                MessageBox.Show("请选择传感器序号");
                return;
            }

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;

            byte[] content = new byte[7];
            content[0] = (byte)comboboxSensorIdx.SelectedIndex;
            uint idx = Convert.ToUInt32(textRecordCount.Text);
            uint count = idx;

            if (count == 0)
            {
                MessageBox.Show("记录数不能为0");
                return;
            }

            Cursor.Current = Cursors.WaitCursor;
            try
            {
                while (count > 0)
                {
                    uint queryRecordCount = 0;
                    if (count > 25)
                    {
                        queryRecordCount = 25;
                    }
                    else
                    {
                        queryRecordCount = count;
                    }



                    content[1] = (byte)((idx >> 24) & 0xFF);
                    content[2] = (byte)((idx >> 16) & 0xFF);
                    content[3] = (byte)((idx >> 8) & 0xFF);
                    content[4] = (byte)(idx & 0xFF);
                    content[5] = (byte)((queryRecordCount >> 8) & 0xFF);
                    content[6] = (byte)(queryRecordCount & 0xFF);

                    byte[] queryMsg = LoadMsg(
                        CMD_QUERY_COLLECTION_DATA,
                        targetAddress,
                        sourceAddr,
                        content,
                        (byte)content.Length);

                    PrintHex(queryMsg);
                    SendHex(queryMsg);

                    byte cmd = 0;
                    uint responseTargetAddress = 0;
                    uint responseSourceAddress = 0;

                    byte[] responseContent;

                    if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent, 10000))
                    {
                        if (cmd == CMD_QUERY_COLLECTION_DATA_REPLY)
                        {
                            PrintHex(responseContent);

                            if (responseContent == null)
                            {
                                MessageBox.Show("传感器没有数据记录");
                                return;
                            }

                            // 处理查询传感器数据的响应
                            if (responseContent.Length >= 10)
                            {
                                int recordCount = responseContent.Length / 10;
                                for (int i = 0; i < recordCount; i++)
                                {
                                    uint timestamp = (uint)(responseContent[i * 10] << 24 |
                                                        responseContent[i * 10 + 1] << 16 |
                                                        responseContent[i * 10 + 2] << 8 |
                                                        responseContent[i * 10 + 3]);

                                    ushort tempture = (ushort)(responseContent[i * 10 + 4] << 8 | responseContent[i * 10 + 5]);
                                    ushort humidity = (ushort)(responseContent[i * 10 + 6] << 8 | responseContent[i * 10 + 7]);
                                    ushort conductivity = (ushort)(responseContent[i * 10 + 8] << 8 | responseContent[i * 10 + 9]);

                                    string data = $"{DateTimeOffset.FromUnixTimeSeconds(timestamp).ToString("yyyy-MM-dd HH:mm:ss")}, " +
                                                $"温度: {tempture / 10.0}°C, " +
                                                $"湿度: {humidity / 10.0}%, " +
                                                $"电导率: {conductivity / 1.0}μS/cm";

                                    Console.WriteLine(data);

                                    string timeStr = DateTimeOffset.FromUnixTimeSeconds(timestamp).ToString("yyyy-MM-dd HH:mm:ss");
                                    string tempStr = (tempture / 10.0).ToString("F1");
                                    string humiStr = (humidity / 10.0).ToString("F1");
                                    string condStr = (conductivity / 1.0).ToString("F2");
                                    string id = (comboboxSensorIdx.SelectedIndex + 1).ToString();
                                    // 添加到DataGridView
                                    if (dataGridViewSensor.InvokeRequired)
                                    {
                                        dataGridViewSensor.Invoke(new Action(() =>
                                        {
                                            dataGridViewSensor.Rows.Add(timeStr, id, tempStr, humiStr, condStr);
                                        }));
                                    }
                                    else
                                    {
                                        dataGridViewSensor.Rows.Add(timeStr, id, tempStr, humiStr, condStr);
                                    }

                                }
                            }
                            else
                            {
                                MessageBox.Show("响应内容格式错误");
                            }
                        }
                        else
                        {
                            MessageBox.Show($"收到未知命令: {cmd:X2}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("接收数据失败");
                    }

                    count -= queryRecordCount;
                    idx -= queryRecordCount;
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void btQueryAllData_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            Cursor.Current = Cursors.WaitCursor;
            try
            {
                uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
                uint sourceAddress = ADDR_UPPER;
                byte[] queryDeviceInfo = LoadMsg(
                    CMD_QUERY_DEVICE_INFO,
                    targetAddress,
                    sourceAddress,
                    null,
                    0);
                PrintHex(queryDeviceInfo);
                SendHex(queryDeviceInfo);
                byte cmd = 0;
                uint receivedTargetAddress = 0;
                uint receivedSourceAddress = 0;
                byte[] content;

                byte sensor_num = 0;
                if (ReceiveMsg(out receivedTargetAddress, out receivedSourceAddress, out cmd, out content))
                {
                    if (cmd == CMD_QUERY_DEVICE_INFO_REPLY)
                    {
                        // 处理查询设备信息的响应
                        if (content != null && content.Length >= 4)
                        {
                            PrintHex(content);

                            sensor_num = content[8];
                        }
                    }
                }

                Console.WriteLine($"传感器数量: {sensor_num}");

                for (int current_sensor_id = 0; current_sensor_id < sensor_num; current_sensor_id++)
                {
                    uint current_sensor_count = 0;

                    byte[] query_time_count_content = new byte[1];
                    query_time_count_content[0] = (byte)current_sensor_id;

                    byte[] queryMsg = LoadMsg(
                        CMD_QUERY_COLLECTION_START_TIME,
                        targetAddress,
                        sourceAddress,
                        query_time_count_content,
                        (byte)query_time_count_content.Length);

                    PrintHex(queryMsg);
                    SendHex(queryMsg);

                    cmd = 0;
                    uint responseTargetAddress = 0;
                    uint responseSourceAddress = 0;

                    byte[] responseContent;

                    if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
                    {
                        if (cmd == CMD_QUERY_COLLECTION_START_TIME_REPLY)
                        {
                            PrintHex(responseContent);

                            // 处理查询启动时间计数的响应
                            if (responseContent.Length == 8)
                            {
                                current_sensor_count = (uint)(responseContent[4] << 24 | responseContent[5] << 16 | responseContent[6] << 8 | responseContent[7]);
                            }
                            else
                            {
                                MessageBox.Show("响应内容格式错误");
                            }
                        }
                        else
                        {
                            MessageBox.Show($"收到未知命令: {cmd:X2}");
                        }
                    }

                    Console.WriteLine($"传感器 {current_sensor_id + 1} 的数据记录数: {current_sensor_count}");



                    if (current_sensor_count > 0)
                    {
                        // 发送查询传感器数据的命令
                        byte[] query_data_content = new byte[7];
                        query_data_content[0] = (byte)current_sensor_id;
                        uint idx = current_sensor_count;
                        uint count = current_sensor_count;

                        while (count > 0)
                        {
                            uint queryRecordCount = 0;
                            if (count > 25)
                            {
                                queryRecordCount = 25;
                            }
                            else
                            {
                                queryRecordCount = count;
                            }

                            query_data_content[0] = (byte)current_sensor_id;
                            query_data_content[1] = (byte)((idx >> 24) & 0xFF);
                            query_data_content[2] = (byte)((idx >> 16) & 0xFF);
                            query_data_content[3] = (byte)((idx >> 8) & 0xFF);
                            query_data_content[4] = (byte)(idx & 0xFF);
                            query_data_content[5] = (byte)((queryRecordCount >> 8) & 0xFF);
                            query_data_content[6] = (byte)(queryRecordCount & 0xFF);

                            queryMsg = LoadMsg(
                                CMD_QUERY_COLLECTION_DATA,
                                targetAddress,
                                sourceAddress,
                                query_data_content,
                                (byte)query_data_content.Length);

                            PrintHex(queryMsg);
                            SendHex(queryMsg);

                            byte cmd2 = 0;
                            uint responseTargetAddress2 = 0;
                            uint responseSourceAddress2 = 0;


                            if (ReceiveMsg(out responseTargetAddress2, out responseSourceAddress2, out cmd2, out responseContent, 10000))
                            {
                                if (cmd2 == CMD_QUERY_COLLECTION_DATA_REPLY)
                                {
                                    PrintHex(responseContent);

                                    // 处理查询传感器数据的响应
                                    if (responseContent.Length >= 10)
                                    {
                                        int recordCount = responseContent.Length / 10;
                                        for (int i = 0; i < recordCount; i++)
                                        {
                                            uint timestamp = (uint)(responseContent[i * 10] << 24 |
                                                                responseContent[i * 10 + 1] << 16 |
                                                                responseContent[i * 10 + 2] << 8 |
                                                                responseContent[i * 10 + 3]);
                                            ushort tempture = (ushort)(responseContent[i * 10 + 4] << 8 | responseContent[i * 10 + 5]);
                                            ushort humidity = (ushort)(responseContent[i * 10 + 6] << 8 | responseContent[i * 10 + 7]);
                                            ushort conductivity = (ushort)(responseContent[i * 10 + 8] << 8 | responseContent[i * 10 + 9]);


                                            string timeStr = DateTimeOffset.FromUnixTimeSeconds(timestamp).ToString("yyyy-MM-dd HH:mm:ss");
                                            string tempStr = (tempture / 10.0).ToString("F1");
                                            string humiStr = (humidity / 10.0).ToString("F1");
                                            string condStr = (conductivity / 1.0).ToString("F2");
                                            string id = (current_sensor_id + 1).ToString();
                                            // 添加到DataGridView
                                            if (dataGridViewSensor.InvokeRequired)
                                            {
                                                dataGridViewSensor.Invoke(new Action(() =>
                                                {
                                                    dataGridViewSensor.Rows.Add(timeStr, id, tempStr, humiStr, condStr);
                                                }));
                                            }
                                            else
                                            {
                                                dataGridViewSensor.Rows.Add(timeStr, id, tempStr, humiStr, condStr);
                                            }

                                            dataGridViewSensor.PerformLayout();
                                            dataGridViewSensor.Controls[1].Enabled = true;
                                            dataGridViewSensor.Controls[1].Visible = true;
                                        }
                                    }
                                }
                            }

                            count -= queryRecordCount;
                            idx -= queryRecordCount;

                        }
                    }



                }

            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void btClearTableData_Click(object sender, EventArgs e)
        {
            if (dataGridViewSensor.InvokeRequired)
            {
                dataGridViewSensor.Invoke(new Action(() => dataGridViewSensor.Rows.Clear()));
            }
            else
            {
                dataGridViewSensor.Rows.Clear();
            }
        }

        private void btQueryRealTimeData_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;


            byte[] queryMsg = LoadMsg(
                CMD_QUERY_REAL_TIME_SENSOR_DATA,
                targetAddress,
                sourceAddr,
                Array.Empty<byte>(),
                0);

            PrintHex(queryMsg);
            SendHex(queryMsg);

            byte cmd = 0;
            uint responseTargetAddress = 0;
            uint responseSourceAddress = 0;

            byte[] responseContent;

            if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_QUERY_REAL_TIME_SENSOR_DATA_REPLY)
                {
                    PrintHex(responseContent);

                    if (responseContent == null || responseContent.Length < 12)
                    {
                        return;
                    }

                    int sensor_count = responseContent.Length / 12;

                    for (int i = 0; i < sensor_count; i++)
                    {
                        byte sensor_id = responseContent[i * 12];
                        uint timestamp = (uint)(responseContent[i * 12 + 1] << 24 |
                                                responseContent[i * 12 + 2] << 16 |
                                                responseContent[i * 12 + 3] << 8 |
                                                responseContent[i * 12 + 4]);
                        ushort tempture = (ushort)(responseContent[i * 12 + 5] << 8 | responseContent[i * 12 + 6]);
                        ushort humidity = (ushort)(responseContent[i * 12 + 7] << 8 | responseContent[i * 12 + 8]);
                        ushort conductivity = (ushort)(responseContent[i * 12 + 9] << 8 | responseContent[i * 12 + 10]);
                        // 处理每个传感器的数据

                        string timeStr = DateTimeOffset.FromUnixTimeSeconds(timestamp).ToString("yyyy-MM-dd HH:mm:ss");
                        string tempStr = (tempture / 10.0).ToString("F1");
                        string humidityStr = (humidity / 10.0).ToString("F1");
                        string conductivityStr = (conductivity / 1.0).ToString("F1");

                        TextBox timeTextBox = this.Controls.Find("textBoxTime" + (sensor_id + 1), true).FirstOrDefault() as TextBox;
                        if (timeTextBox != null)
                        {
                            timeTextBox.Text = timeStr;
                        }

                        TextBox tempTextBox = this.Controls.Find("textBoxTemp" + (sensor_id + 1), true).FirstOrDefault() as TextBox;
                        if (tempTextBox != null)
                        {
                            tempTextBox.Text = tempStr;
                        }

                        TextBox humidityTextBox = this.Controls.Find("textBoxHumidity" + (sensor_id + 1), true).FirstOrDefault() as TextBox;
                        if (humidityTextBox != null)
                        {
                            humidityTextBox.Text = humidityStr;
                        }

                        TextBox conductivityTextBox = this.Controls.Find("textBoxConductivity" + (sensor_id + 1), true).FirstOrDefault() as TextBox;
                        if (conductivityTextBox != null)
                        {
                            conductivityTextBox.Text = conductivityStr;
                        }


                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void btQuerySensorModel_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;


            byte[] queryMsg = LoadMsg(
                CMD_QUERY_SENSOR_MODEL,
                targetAddress,
                sourceAddr,
                Array.Empty<byte>(),
                0);

            PrintHex(queryMsg);
            SendHex(queryMsg);

            byte cmd = 0;
            uint responseTargetAddress = 0;
            uint responseSourceAddress = 0;

            byte[] responseContent;

            if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_QUERY_SENSOR_MODEL_REPLY)
                {
                    PrintHex(responseContent);

                    for (int i = 0; i < responseContent.Length; i++)
                    {
                        ComboBox modelComboBox = this.Controls.Find("comboBoxSensorModel" + (i + 1), true).FirstOrDefault() as ComboBox;
                        if (modelComboBox != null)
                        {
                            if (responseContent[i] == 0)
                            {
                                modelComboBox.SelectedIndex = 0; // 假设0表示无传感器
                            }
                            else if (responseContent[i] == 1)
                            {
                                modelComboBox.SelectedIndex = 1; // PL30
                            }
                            else
                            {
                                modelComboBox.SelectedIndex = -1; // 未知传感器类型
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void btSetSensorModel_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            byte[] content = new byte[10];

            for (int i = 0; i < 10; i++)
            {
                ComboBox modelComboBox = this.Controls.Find("comboBoxSensorModel" + (i + 1), true).FirstOrDefault() as ComboBox;
                if (modelComboBox != null)
                {
                    if (modelComboBox.SelectedIndex == 0)
                    {
                        content[i] = 0; // 无传感器
                    }
                    else if (modelComboBox.SelectedIndex == 1)
                    {
                        content[i] = 1; // PL30
                    }
                    else
                    {
                        content[i] = 0; // 未知传感器类型，默认为0
                    }
                }
            }

            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;


            byte[] queryMsg = LoadMsg(
                CMD_SET_SENSOR_MODEL,
                targetAddress,
                sourceAddr,
                content,
                (byte)content.Length);

            PrintHex(queryMsg);
            SendHex(queryMsg);

            byte cmd = 0;
            uint responseTargetAddress = 0;
            uint responseSourceAddress = 0;

            byte[] responseContent;

            if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
            {
                if (cmd == CMD_SET_SENSOR_MODEL_REPLY)
                {
                    PrintHex(responseContent);

                    if (responseContent[0] == 1)
                    {
                        MessageBox.Show("传感器型号设置成功");
                    }
                    else
                    {
                        MessageBox.Show("传感器型号设置失败");
                    }
                }
                else
                {
                    MessageBox.Show($"收到未知命令: {cmd:X2}");
                }
            }
        }

        private void btDevReset_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }


            uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
            uint sourceAddr = ADDR_UPPER;


            byte[] queryMsg = LoadMsg(
                CMD_RESET_DEVICE,
                targetAddress,
                sourceAddr,
                Array.Empty<byte>(),
                (byte)Array.Empty<byte>().Length);

            PrintHex(queryMsg);
            SendHex(queryMsg);

            MessageBox.Show("设备复位命令已发送。");

        }

        private void btRestart_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }

            if (comboBoxDevAdd.Items.Count == 0)
            {
                MessageBox.Show("请先查询设备地址");
                return;
            }

            if (comboBoxDevAdd.SelectedIndex < 0)
            {
                MessageBox.Show("请选择设备地址");
                return;
            }

            var result = MessageBox.Show(
                        $"此操作将清空设备中所有传感器数据，是否继续？",
                        "危险操作确认",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

            Cursor.Current = Cursors.WaitCursor;

            try
            {
                if (result == DialogResult.Yes)
                {
                    // 执行清空数据命令


                    uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
                    uint sourceAddr = ADDR_UPPER;


                    byte[] queryMsg = LoadMsg(
                        CMD_COLLECTION_RESTART,
                        targetAddress,
                        sourceAddr,
                        Array.Empty<byte>(),
                        0);

                    PrintHex(queryMsg);
                    SendHex(queryMsg);


                    byte cmd = 0;
                    uint responseTargetAddress = 0;
                    uint responseSourceAddress = 0;

                    byte[] responseContent;

                    if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
                    {
                        if (cmd == CMD_COLLECTION_RESTART_REPLY)
                        {
                            PrintHex(responseContent);

                            if (responseContent[0] == 1)
                            {
                                MessageBox.Show($"传感器数据清空成功");
                            }
                            else
                            {
                                MessageBox.Show($"传感器数据清空失败");
                            }
                        }
                        else
                        {
                            MessageBox.Show($"收到未知命令: {cmd:X2}");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("操作已取消。");
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void btClearSensorData_Click(object sender, EventArgs e)
        {
            byte sensor_id = (byte)comboboxSensorIdx.SelectedIndex;
            if (sensor_id < 0)
            {
                MessageBox.Show("请选择传感器序号");
                return;
            }

            var result = MessageBox.Show(
                        $"此操作将清空设备中传感器“{sensor_id + 1}”数据，是否继续？",
                        "危险操作确认",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                // 执行清空数据命令
                if (!serialPort.IsOpen)
                {
                    MessageBox.Show("请先打开串口");
                    return;
                }

                if (comboBoxDevAdd.Items.Count == 0)
                {
                    MessageBox.Show("请先查询设备地址");
                    return;
                }

                if (comboBoxDevAdd.SelectedIndex < 0)
                {
                    MessageBox.Show("请选择设备地址");
                    return;
                }

                uint targetAddress = Convert.ToUInt32(comboBoxDevAdd.Text, 16);
                uint sourceAddr = ADDR_UPPER;

                byte[] content = new byte[1];
                content[0] = sensor_id; // 传感器ID从0开始

                byte[] queryMsg = LoadMsg(
                    CMD_CLEAR_COLLECTION_DATA,
                    targetAddress,
                    sourceAddr,
                    content,
                    (byte)content.Length);

                PrintHex(queryMsg);
                SendHex(queryMsg);


                byte cmd = 0;
                uint responseTargetAddress = 0;
                uint responseSourceAddress = 0;

                byte[] responseContent;

                if (ReceiveMsg(out responseTargetAddress, out responseSourceAddress, out cmd, out responseContent))
                {
                    if (cmd == CMD_CLEAR_COLLECTION_DATA_REPLY)
                    {
                        PrintHex(responseContent);

                        if (responseContent[0] == 1)
                        {
                            MessageBox.Show($"传感器{sensor_id + 1}数据清空成功");
                        }
                        else
                        {
                            MessageBox.Show($"传感器{sensor_id + 1}数据清空失败");
                        }
                    }
                    else
                    {
                        MessageBox.Show($"收到未知命令: {cmd:X2}");
                    }
                }
            }
            else
            {
                MessageBox.Show("操作已取消。");
            }
        }

        private void AddLogToTextBox(string log)
        {
            if (textBoxLog.InvokeRequired)
            {
                textBoxLog.Invoke(new Action(() => textBoxLog.AppendText(log + Environment.NewLine)));
            }
            else
            {
                textBoxLog.AppendText(log + Environment.NewLine);
            }
        }

        private void btClearLog_Click(object sender, EventArgs e)
        {
            if (textBoxLog.InvokeRequired)
            {
                textBoxLog.Invoke(new Action(() => textBoxLog.Clear()));
            }
            else
            {
                textBoxLog.Clear();
            }
        }

        private void btSaveLog_Click(object sender, EventArgs e)
        {
            if (textBoxLog.InvokeRequired)
            {
                textBoxLog.Invoke(new Action(() => SaveLogToFile()));
            }
            else
            {
                SaveLogToFile();
            }
        }
        
        private void SaveLogToFile()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";
                saveFileDialog.DefaultExt = "txt";
                saveFileDialog.AddExtension = true;
                saveFileDialog.FileName = $"log-{DateTime.Now:yyyyMMdd_HHmmss}.txt";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        System.IO.File.WriteAllText(saveFileDialog.FileName, textBoxLog.Text);
                        MessageBox.Show("日志已保存到 " + saveFileDialog.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("保存日志失败: " + ex.Message);
                    }
                }
            }
        }
    }
}
