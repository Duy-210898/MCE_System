using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows.Forms;
using EasyModbus;
using System.Drawing;
using OfficeOpenXml;
using System.Threading;
using System.Linq;
using Microsoft.Win32;

namespace Measure_Energy_Consumption
{
    public partial class frmMain : XtraForm
    {
        private ModbusClient modbusClient;
        private DataTable machine1DataTable;
        private DataTable machine2DataTable;
        private bool shouldContinueIp1 = true;
        private bool shouldContinueIp2 = true;

        private DataTable cabinet1DataTable;
        private DataTable cabinet2DataTable;

        private NotifyIcon trayIcon;

        public frmMain()
        {
            InitializeComponent();

            SetStyle(ControlStyles.DoubleBuffer | ControlStyles.OptimizedDoubleBuffer, true);

            dgvHourlyIp1.DoubleBuffered(true);
            dgvHourlyIp2.DoubleBuffered(true);
            dgvDataIp1.DoubleBuffered(true);
            dgvDataIp2.DoubleBuffered(true);

            EnableDoubleBuffering(this);

            SetDefaultTime();
            InitializeDataTables();

            modbusClient = new ModbusClient();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            numTimeRecordIp1.Value = 1;
            numTimeRecordIp2.Value = 1;

            InitializeNotifyIcon();
        }

        private void InitializeNotifyIcon()
        {
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Text = "Measure Energy Consumption";
            trayIcon.Icon = Icon.FromHandle(Properties.Resources.electric_power_meter_energy_electricity_counter_vector_48294335_removebg_preview.GetHicon());
            trayIcon.Visible = true;

            // Tạo menu context cho NotifyIcon
            ContextMenu contextMenu = new ContextMenu();
            contextMenu.MenuItems.Add("Show Application", (sender, e) =>
            {
                Show();
                WindowState = FormWindowState.Normal;
            });
            contextMenu.MenuItems.Add("Exit", (sender, e) =>
            {
                Application.Exit();
            });
            trayIcon.ContextMenu = contextMenu;

            trayIcon.DoubleClick += trayIcon_DoubleClick;
        }

        // Xử lý sự kiện double-click trên NotifyIcon để hiện lại cửa sổ chính
        private void trayIcon_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        /// <summary>
        /// Override phương thức OnFormClosing để chuyển ứng dụng vào background khi bị đóng
        /// </summary>
        /// <param name="e"></param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; 
                Hide(); 
            }
        }

        /// <summary>
        /// Khởi tạo DataTables
        /// </summary>
        private void InitializeDataTables()
        {
            machine1DataTable = new DataTable();
            machine2DataTable = new DataTable();

            cabinet1DataTable = new DataTable();
            cabinet2DataTable = new DataTable();

            dgvHourlyIp1.RowHeadersVisible = false;
            dgvHourlyIp2.RowHeadersVisible = false;

            dgvDataIp1.RowHeadersVisible = false;
            dgvDataIp2.RowHeadersVisible = false;
        }

        /// <summary>
        /// Giảm flickering cho các controls
        /// </summary>
        /// <param name="control"></param>
        private void EnableDoubleBuffering(Control control)
        {
            typeof(Control).InvokeMember("DoubleBuffered", System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, null, control, new object[] { true });

            // Duyệt qua tất cả các control con và gọi đệ quy để bật double buffering
            foreach (Control childControl in control.Controls)
            {
                EnableDoubleBuffering(childControl);
            }
        }

        /// <summary>
        /// Biến kiểm tra giờ hiện tại có nằm ở các mốc cần tính toán không.
        /// </summary>
        /// <param name="currentTime"></param>
        /// <returns></returns>
        public bool IsHourToCalculate(DateTime currentTime)
        {
            int hour = currentTime.Hour;
            int minute = currentTime.Minute;

            if (minute == 30 && hour >= 8 && hour <= 18)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Nếu là giờ cần tính toán thì gọi hàm tính toán, nghỉ 1p trước khi thực hiện kiểm tra lại
        /// Cho Cabinet 1
        /// </summary>
        public void CheckAndCalculateHourlyEnergy1()
        {
            while (true)
            {
                DateTime currentTime = DateTime.Now;

                if (IsHourToCalculate(currentTime))
                {
                    CalculateHourlyEnergy(machine1DataTable, dgvHourlyIp1, "Cabinet 1");
                }

                Thread.Sleep(60000);
            }
        }

        /// <summary>
        /// Nếu là giờ cần tính toán thì gọi hàm tính toán, nghỉ 1p trước khi thực hiện kiểm tra lại
        /// Cho Cabinet 2
        /// </summary>
        public void CheckAndCalculateHourlyEnergy2()
        {
            while (true)
            {
                DateTime currentTime = DateTime.Now;

                if (IsHourToCalculate(currentTime))
                {
                    CalculateHourlyEnergy(machine2DataTable, dgvHourlyIp2, "Cabinet 2");
                }

                Thread.Sleep(60000);
            }
        }

        private void StopPeriodicIp1()
        {
            shouldContinueIp1 = false;
        }

        private void StopPeriodicIp2()
        {
            shouldContinueIp2 = false;
        }

        /// <summary>
        /// 
        /// </summary>
        private void SetDefaultTime()
        {
            dtpEndTimeIp1.ShowUpDown = true;
            dtpEndTimeIp2.ShowUpDown = true;
            dtpStartTimeIp1.ShowUpDown = true;
            dtpStartTimeIp2.ShowUpDown = true;

            dtpStartTimeIp1.Value = DateTime.Today.AddHours(7).AddMinutes(30);
            dtpEndTimeIp1.Value = DateTime.Today.AddHours(17);

            dtpStartTimeIp2.Value = DateTime.Today.AddHours(7).AddMinutes(30);
            dtpEndTimeIp2.Value = DateTime.Today.AddHours(17);
        }

        /// <summary>
        /// Từ điển lưu các giá trị trạng thái
        /// </summary>
        private Dictionary<string, int> timeDelays = new Dictionary<string, int>();
        private readonly Dictionary<string, ModbusClient> connectedDevices = new Dictionary<string, ModbusClient>();
        private readonly Dictionary<string, bool> isConnected = new Dictionary<string, bool>();


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void cbxListIp1_Click(object sender, EventArgs e)
        {
            await UpdateDeltaDevicesComboBoxAsync(cbxListIp1, "10.30.4", "192.168.0", "192.168.1", "192.168.2", "192.168.3");
        }

        private async void cbxListIp2_Click(object sender, EventArgs e)
        {
            await UpdateDeltaDevicesComboBoxAsync(cbxListIp2, "10.30.4", "192.168.0", "192.168.1", "192.168.2", "192.168.3");
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDisconnectIp1_Click(object sender, EventArgs e)
        {
            if (lblStatusIp1.Text == "Kết nối thành công")
            {
                DisconnectDevice(cbxListIp1, lblStatusIp1, lblDataIp1, numTimeRecordIp1, ref shouldContinueIp1);
            }
        }

        private void btnDisconnectIp2_Click(object sender, EventArgs e)
        {
            if (lblStatusIp2.Text == "Kết nối thành công")
            {
                DisconnectDevice(cbxListIp2, lblStatusIp2, lblDataIp2, numTimeRecordIp2, ref shouldContinueIp2);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void btnConnectIp1_Click(object sender, EventArgs e)
        {
            if (lblStatusIp1.Text != "Kết nối thành công")
            {
                await ConnectToDeviceAsync(cbxListIp1, lblDataIp1, lblStatusIp1, dgvHourlyIp1);
            }
        }

        private async void btnConnectIp2_Click(object sender, EventArgs e)
        {
            if (lblStatusIp2.Text != "Kết nối thành công")
            {
                await ConnectToDeviceAsync(cbxListIp2, lblDataIp2, lblStatusIp2, dgvHourlyIp2);
            }
        }

        /// <summary>
        /// Tạo datatable tương ứng với từng tủ
        /// </summary>
        /// <param name="comboBox"></param>
        private void CreateCabinetDataTable(System.Windows.Forms.ComboBox comboBox)
        {
            if (comboBox == cbxListIp1)
            {
                CreateCabinet1DataTable();
                dgvHourlyIp1.DataSource = CreateCabinet1DataTable();
            }
            else if (comboBox == cbxListIp2)
            {
                CreateCabinet2DataTable();
                dgvHourlyIp2.DataSource = CreateCabinet2DataTable();
            }
        }

        /// <summary>
        /// Tạo dataTable tủ 1
        /// </summary>
        /// <returns></returns>
        private DataTable CreateCabinet1DataTable()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Thời gian", typeof(string));
            dataTable.Columns.Add("Tiêu thụ - Máy 1\n(Kwh)", typeof(float));
            dataTable.Columns.Add("Tiêu thụ - Máy 2\n(Kwh)", typeof(float));
            return dataTable;
        }

        /// <summary>
        /// Tạo dataTable tủ 2
        /// </summary>
        /// <returns></returns>
        private DataTable CreateCabinet2DataTable()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Thời gian", typeof(string));
            dataTable.Columns.Add("Tiêu thụ - Máy 1\n(Kwh)", typeof(float));
            dataTable.Columns.Add("Tiêu thụ - Máy 2\n(Kwh)", typeof(float));
            return dataTable;
        }

        /// <summary>
        /// Lấy giá trị và tính toán điện năng tiêu thụ theo giờ
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="dataGridView"></param>
        /// <param name="cabinet"></param>
        public static void CalculateHourlyEnergy(DataTable dataTable, DataGridView dataGridView, string cabinet)
        {
            DateTime currentTime = DateTime.Now;
            DateTime startTime = currentTime.AddHours(-1);

            string register;
             
            if (cabinet == "Cabinet 1")
            {
                register = "Register_6";
            }
            else if (cabinet == "Cabinet 2")
            {
                register = "Register_16";
            }
            else
            {
                return;
            }

            // Lấy giá trị của thanh ghi số từ dòng đầu tiên tìm thấy sau startTime
            float startValueRegister1 = GetStartValue(dataTable, startTime, currentTime, "Register_0");
            float startValueRegister2 = GetStartValue(dataTable, startTime, currentTime, register);

            // Lấy giá trị của thanh ghi số hiện tại
            float endValueRegister0 = GetEndValue(dataTable, currentTime, "Register_0");
            float endValueRegister6 = GetEndValue(dataTable, currentTime, register);

            // Làm tròn kết quả về 2 số phần thập phân
            startValueRegister1 = (float)Math.Round(startValueRegister1, 2);
            startValueRegister2 = (float)Math.Round(startValueRegister2, 2);
            endValueRegister0 = (float)Math.Round(endValueRegister0, 2);
            endValueRegister6 = (float)Math.Round(endValueRegister6, 2);

            // Tính toán totalEnergy1 và totalEnergy2
            float totalEnergy1 = (float)Math.Round(endValueRegister0 - startValueRegister1, 2);
            float totalEnergy2 = (float)Math.Round(endValueRegister6 - startValueRegister2, 2);

            // Tạo chuỗi thời gian cho dòng hiện tại
            string timeRange = $"{startTime.ToString("HH:mm")} - {currentTime.ToString("HH:mm")}";

            // Sử dụng phương thức Invoke để thực hiện thay đổi trên UI thread
            dataGridView.Invoke((MethodInvoker)delegate
            {
                // Thêm dữ liệu vào dòng tương ứng trong dataGridView
                ((DataTable)dataGridView.DataSource).Rows.Add(timeRange, totalEnergy1, totalEnergy2);

                // Làm mới DataGridView để hiển thị dữ liệu mới
                dataGridView.Refresh();
            });

            // Gọi phương thức UpdateExcelWithHourlyData để cập nhật dữ liệu vào tệp Excel
            ExcelUpdate.UpdateExcelWithHourlyData(totalEnergy1, totalEnergy2, startTime, currentTime, cabinet);
        }


        /// <summary>
        /// Lấy giá trị bắt đầu của thanh ghi
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private static float GetStartValue(DataTable dataTable, DateTime startTime, DateTime endTime, string columnName)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                if (DateTime.TryParse(row["Thời gian"].ToString(), out DateTime timeStamp))
                {
                    if (timeStamp >= startTime && timeStamp <= endTime)
                    {
                        return Convert.ToSingle(row[columnName]);
                    }
                }
            }
            return 0;
        }


        /// <summary>
        /// Lấy giá trị kết thúc của thanh ghi
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="currentTime"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private static float GetEndValue(DataTable dataTable, DateTime currentTime, string columnName)
        {
            DataRow row = dataTable.AsEnumerable()
                                   .OrderByDescending(r => DateTime.TryParse(r["Thời gian"].ToString(), out DateTime timeStamp) ? timeStamp : DateTime.MinValue)
                                   .FirstOrDefault();

            if (row != null)
            {
                return Convert.ToSingle(row[columnName]);
            }
            return 0;
        }

        /// <summary>
        /// Biên theo dõi trạng thái hiển thị của messagebox
        /// </summary>
        private bool isErrorMessageBoxShown = false;
        private bool isMessageBoxShown = false;


        /// <summary>
        /// Phương thức kết nối đến thiết bị (thực hiện bất đồng bộ)
        /// </summary>
        /// <param name="comboBox"></param>
        /// <param name="dataLabel"></param>
        /// <param name="statusLabel"></param>
        /// <param name="numericUpDown"></param>
        /// <param name="dataGridView"></param>
        /// <returns></returns>
        private async Task<bool> ConnectToDeviceAsync(System.Windows.Forms.ComboBox comboBox, Label dataLabel, Label statusLabel, DataGridView dataGridView)
        {
            if (comboBox.SelectedItem != null && !string.IsNullOrEmpty(comboBox.SelectedItem.ToString()))
            {
                string ipAddress = comboBox.SelectedItem.ToString();

                if (isConnected.ContainsKey(ipAddress) && isConnected[ipAddress])
                {
                    ShowConnectionMessageBox();
                    return true;
                }

                string cabinet = GetCabinetName(dataLabel);

                DisplayDeviceInformation(dataLabel, cabinet, ipAddress);

                try
                {
                    await ConnectToModbusDevice(ipAddress);

                    UpdateConnectionStatus(statusLabel);

                    comboBox.Enabled = false;

                    CreateCabinetDataTable(comboBox);

                    FormatHourlyDataGridView(dataGridView);

                    await StartPeriodicDataReading(comboBox);

                    return true;
                }
                catch (OperationCanceledException)
                {
                    return false;
                }
                catch (SocketException ex)
                {
                    DisplayConnectionErrorMessageBox(ex, statusLabel);
                }
                catch (TimeoutException ex)
                {
                    DisplayTimeoutErrorMessageBox(ex, statusLabel);
                }
                catch (Exception ex)
                {
                    DisplayGenericErrorMessageBox(ex, statusLabel);
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn địa chỉ IP.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return false;
        }


        /// <summary>
        /// 
        /// </summary>
        private void ShowConnectionMessageBox()
        {
            if (!isMessageBoxShown)
            {
                Thread messageThread = new Thread(() =>
                {
                    MessageBox.Show("Đang kết nối đến địa chỉ IP này.\nVui lòng chọn IP khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    isMessageBoxShown = true;
                });
                messageThread.Start();
            }
        }

        /// <summary>
        /// Lấy tên tủ hiện tại
        /// </summary>
        /// <param name="dataLabel"></param>
        /// <returns></returns>
        private string GetCabinetName(Label dataLabel)
        {
            string labelName = dataLabel.Name;
            char labelNumber = labelName[labelName.Length - 1];

            string cabinet;
            if (labelNumber == '1')
            {
                cabinet = "Cabinet 1";
            }
            else if (labelNumber == '2')
            {
                cabinet = "Cabinet 2";
            }
            else
            {
                cabinet = "Tủ không xác định";
            }

            return cabinet;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataLabel"></param>
        /// <param name="cabinet"></param>
        /// <param name="ipAddress"></param>
        private void DisplayDeviceInformation(Label dataLabel, string cabinet, string ipAddress)
        {
            dataLabel.Text = $"{cabinet}: {ipAddress}";
            dataLabel.ForeColor = Color.Green;
        }

        /// <summary>
        /// Kết nối đến thiết bị
        /// </summary>
        /// <param name="ipAddress"></param>
        /// <returns></returns>
        private async Task ConnectToModbusDevice(string ipAddress)
        {
            ModbusClient modbusClient = new ModbusClient(ipAddress, 502)
            {
                ConnectionTimeout = 15000,
            };

            await Task.Run(() => modbusClient.Connect());
            connectedDevices[ipAddress] = modbusClient;
            isConnected[ipAddress] = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="statusLabel"></param>
        private void UpdateConnectionStatus(Label statusLabel)
        {
            statusLabel.Text = "Kết nối thành công";
            statusLabel.ForeColor = Color.Green;
        }

        /// <summary>
        /// Bắt đầu cập nhật dữ liệu định kì
        /// </summary>
        /// <param name="comboBox"></param>
        /// <returns></returns>
        private async Task StartPeriodicDataReading(System.Windows.Forms.ComboBox comboBox)
        {
            if (comboBox == cbxListIp1)
            {
                int timeDelay = (int)numTimeRecordIp1.Value;
                await StartPeriodicIp1(timeDelay);
            }
            else if (comboBox == cbxListIp2)
            {
                int timeDelay = (int)numTimeRecordIp2.Value;
                await StartPeriodicIp2(timeDelay);
            }
        }

        /// <summary>
        /// Hiển thị thông báo lỗi liên quan khi kết nối thiết bị
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="statusLabel"></param>
        private void DisplayConnectionErrorMessageBox(SocketException ex, Label statusLabel)
        {
            if (!isErrorMessageBoxShown)
            {
                Thread errorMessageThread = new Thread(() =>
                {
                    MessageBox.Show("Không thể kết nối đến thiết bị: " + ex.Message);
                    statusLabel.ForeColor = Color.Orange;
                    statusLabel.Text = "Kết nối thất bại";
                    isErrorMessageBoxShown = true;
                });
                errorMessageThread.Start();
            }
        }

        /// <summary>
        /// Lỗi TimeOut
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="statusLabel"></param>
        private void DisplayTimeoutErrorMessageBox(TimeoutException ex, Label statusLabel)
        {
            if (!isErrorMessageBoxShown)
            {
                Thread errorMessageThread = new Thread(() =>
                {
                    MessageBox.Show("Hết thời gian chờ kết nối đến thiết bị: " + ex.Message);
                    statusLabel.ForeColor = Color.Orange;
                    statusLabel.Text = "Kết nối thất bại";
                    isErrorMessageBoxShown = true;
                });
                errorMessageThread.Start();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="statusLabel"></param>
        private void DisplayGenericErrorMessageBox(Exception ex, Label statusLabel)
        {
            if (!isErrorMessageBoxShown)
            {
                Thread errorMessageThread = new Thread(() =>
                {
                    MessageBox.Show("Lỗi khi kết nối đến thiết bị: " + ex.Message);
                    statusLabel.ForeColor = Color.Orange;
                    statusLabel.Text = "Kết nối thất bại";
                    isErrorMessageBoxShown = true;
                });
                errorMessageThread.Start();
            }
        }

        /// <summary>
        /// Ngắt kết nối
        /// </summary>
        /// <param name="comboBox"></param>
        /// <param name="statusLabel"></param>
        /// <param name="dataLabel"></param>
        /// <param name="numericUpDown"></param>
        /// <param name="shouldContinue"></param>
        private void DisconnectDevice(System.Windows.Forms.ComboBox comboBox, Label statusLabel, Label dataLabel, NumericUpDown numericUpDown, ref bool shouldContinue)
        {
            if (comboBox.SelectedItem != null)
            {
                string ipAddress = comboBox.SelectedItem.ToString();
                DisconnectFromDevice(ipAddress, statusLabel);

                dataLabel.ForeColor = Color.Red;

                comboBox.Enabled = true;

                if (timeDelays.ContainsKey(ipAddress))
                {
                    int newTimeDelay = (int)numericUpDown.Value;

                    if (timeDelays[ipAddress] != newTimeDelay)
                    {
                        timeDelays[ipAddress] = newTimeDelay;

                        if (comboBox == cbxListIp1)
                        {
                            StopPeriodicIp1();
                        }
                        else if (comboBox == cbxListIp2)
                        {
                            StopPeriodicIp2();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một địa chỉ IP để ngắt kết nối.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ipAddress"></param>
        /// <param name="statusLabel"></param>
        private void DisconnectFromDevice(string ipAddress, Label statusLabel)
        {
            if (!connectedDevices.ContainsKey(ipAddress))
            {
                return;
            }

            ModbusClient modbusClient = connectedDevices[ipAddress];
            if (modbusClient != null && modbusClient.Connected)
            {
                try
                {
                    modbusClient.Disconnect();
                    connectedDevices.Remove(ipAddress);
                    isConnected[ipAddress] = false;
                    statusLabel.Text = "Đã ngắt kết nối";
                    statusLabel.ForeColor = Color.Red;
                }
                catch (Exception ex)
                {
                    statusLabel.Text = "Lỗi khi ngắt kết nối: " + ex.Message;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void ReadAndUpdateIp1Data()
        {
            if (cbxListIp1.InvokeRequired)
            {
                cbxListIp1.Invoke(new MethodInvoker(delegate
                {
                    ReadAndUpdateIp1Data();
                }));
            }
            else
            {
                if (cbxListIp1.SelectedItem != null)
                {
                    string selectedIP = cbxListIp1.SelectedItem.ToString();

                    if (!string.IsNullOrEmpty(selectedIP) && isConnected.ContainsKey(selectedIP) && isConnected[selectedIP])
                    {
                        if (ConnectToDeltaDevice(selectedIP))
                        {
                            int[] registerAddresses = { 0, 2, 4, 6, 8, 10 };
                            float[] data = ReadDataFromRegisters(registerAddresses);
                            string cabinet = "Cabinet 1";

                            for (int i = 0; i < data.Length; i++)
                            {
                                switch (i)
                                {
                                    case 0:
                                    case 6:
                                        data[i] = RoundValue(data[i], 2);
                                        break;
                                    case 2:
                                    case 8:
                                        data[i] = RoundValue(data[i], 1);
                                        break;
                                    case 4:
                                    case 10:
                                        data[i] = RoundValue(data[i], 2);
                                        break;
                                    default:
                                        break;
                                }
                            }

                            UpdateIp1DataTable(data, registerAddresses);

                            ExcelUpdate.UpdateExcelWithNewData(data, cabinet);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <param name="decimalPlaces"></param>
        /// <returns></returns>
        private float RoundValue(float value, int decimalPlaces)
        {
            return (float)Math.Round(value, decimalPlaces, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// 
        /// </summary>
        private void ReadAndUpdateIp2Data()
        {
            if (cbxListIp2.InvokeRequired)
            {
                cbxListIp2.Invoke(new MethodInvoker(delegate
                {
                    ReadAndUpdateIp2Data();
                }));
            }
            else
            {
                if (cbxListIp2.SelectedItem != null)
                {
                    string selectedIP = cbxListIp2.SelectedItem.ToString();

                    if (!string.IsNullOrEmpty(selectedIP) && isConnected.ContainsKey(selectedIP) && isConnected[selectedIP])
                    {
                        if (ConnectToDeltaDevice(selectedIP))
                        {
                            int[] registerAddresses = Enumerable.Range(0, 16).Select(x => x * 2).ToArray();

                            float[] data = ReadDataFromRegisters(registerAddresses);
                            string cabinet = "Cabinet 2";

                            for (int i = 0; i < data.Length; i++)
                            {
                                switch (i)
                                {
                                    case 0:
                                    case 6:
                                        data[i] = RoundValue(data[i], 2);
                                        break;
                                    case 2:
                                    case 8:
                                        data[i] = RoundValue(data[i], 1);
                                        break;
                                    case 4:
                                    case 10:
                                        data[i] = RoundValue(data[i], 2);
                                        break;
                                    default:
                                        break;
                                }
                            }

                            UpdateIp2DataTable(data, registerAddresses);

                            ExcelUpdate.UpdateExcelWithNewData(data, cabinet);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Định dạng dataGridview
        /// </summary>
        /// <param name="dataGridView"></param>
        private void FormatDataGridViewCabinet1(DataGridView dataGridView)
        {
            if (dataGridView == null || dataGridView.Columns.Count == 0)
            {
                return;
            }

            Dictionary<int, string> headerMapping = new Dictionary<int, string>()
    {
        { 0, "Điện năng\nMáy 1\n(Kwh)" },
        { 2, "Điện áp\nMáy 1\n(V)" },
        { 4, "Dòng điện\nMáy 1\n(A)" },
        { 6, "Điện năng\nMáy 2\n(Kwh)" },
        { 8, "Điện áp\nMáy 2\n(V)" },
        { 10, "Dòng điện\nMáy 2\n(A)" }
    };

            int columnIndex = 0;

            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                if (column != null && column.HeaderText.StartsWith("Register_") && column.HeaderText.Length > 9)
                {
                    string registerNumber = column.HeaderText.Substring(9);
                    int registerAddress;
                    if (int.TryParse(registerNumber, out registerAddress) && headerMapping.ContainsKey(registerAddress))
                    {
                        column.HeaderText = headerMapping[registerAddress];
                    }
                }

                if (columnIndex == 1 || columnIndex == 3)
                {
                    column.DefaultCellStyle.BackColor = Color.AliceBlue;
                }
                if (columnIndex == 4 || columnIndex == 6)
                {
                    column.DefaultCellStyle.BackColor = Color.LightYellow;
                }

                if (column.HeaderText == "Thời gian")
                {
                    column.Width = 120;
                }

                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                columnIndex++;
            }
        }

        private void FormatDataGridViewCabinet2(DataGridView dataGridView)
        {
            if (dataGridView == null || dataGridView.Columns.Count == 0)
            {
                return;
            }

            Dictionary<int, string> headerMapping = new Dictionary<int, string>()
    {
        // Mapping giữa số register và tiêu đề cột
        { 0, "Tổng điện năng\nMáy 1\n(Kwh)" },
        { 2, "Điện áp\nPha A\nMáy 1\n(V)" },
        { 4, "Điện áp\nPha B\nMáy 1\n(V)" },
        { 6, "Điện áp\nPha C\nMáy 1\n(V)" },
        { 8, "Dòng điện\nPha A\nMáy 1\n(A)" },
        { 10, "Dòng điện\nPha B\nMáy 1\n(A)" },
        { 12, "Điện áp\nPha C\nMáy 1\n(A)" },
        { 14, "Công suất\nMáy 1\n(Kw)" },
        { 16, "Tổng điện năng\nMáy 2\n(Kwh)" },
        { 18, "Điện áp\nPha A\nMáy 2\n(V)" },
        { 20, "Điện áp\nPha B\nMáy 2\n(V)" },
        { 22, "Điện áp\nPha C\nMáy 2\n(V)" },
        { 24, "Dòng điện\nPha A\nMáy 2\n(A)" },
        { 26, "Dòng điện\nPha B\nMáy 2\n(A)" },
        { 28, "Điện áp\nPha C\nMáy 2\n(A)" },
        { 30, "Công suất\nMáy 2\n(Kw)" }
    };

            int columnIndex = 0;
            Color[] alternatingColors = { Color.AliceBlue, Color.White };
            Color[] alternatingColors2 = { Color.White, Color.LightYellow };
            int colorIndex = 0;

            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                if (column != null && column.HeaderText.StartsWith("Register_") && column.HeaderText.Length > 9)
                {
                    string registerNumber = column.HeaderText.Substring(9);
                    int registerAddress;
                    if (int.TryParse(registerNumber, out registerAddress))
                    {
                        if (headerMapping.ContainsKey(registerAddress))
                        {
                            column.HeaderText = headerMapping[registerAddress];
                        }
                    }
                }

                // Tô màu xen kẽ giữa AliceBlue và LightYellow từ cột 2 đến cột 9
                if (columnIndex >= 1 && columnIndex <= 8)
                {
                    column.DefaultCellStyle.BackColor = alternatingColors[colorIndex % alternatingColors.Length];
                }
                // Tô màu xen kẽ giữa Gray và White từ cột 10 đến cột 17
                else if (columnIndex >= 9 && columnIndex <= 16)
                {
                    column.DefaultCellStyle.BackColor = alternatingColors2[colorIndex % alternatingColors2.Length];
                }

                if (column.HeaderText == "Thời gian")
                {
                    column.Width = 90;
                }

                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                columnIndex++;
                colorIndex++;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataGridView"></param>
        private void FormatHourlyDataGridView(DataGridView dataGridView)
        {
            if (dataGridView == null || dataGridView.Columns.Count == 0)
            {
                return;
            }

            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            dataGridView.RowHeadersVisible = false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="registerAddresses"></param>
        private void UpdateIp1DataTable(float[] data, int[] registerAddresses)
        {
            // Kiểm tra xem cột Thời gian đã tồn tại trong DataTable chưa
            if (!machine1DataTable.Columns.Contains("Thời gian"))
            {
                // Nếu chưa tồn tại, thêm cột mới có tên là "Thời gian" và kiểu dữ liệu là string (định dạng HH:mm:ss)
                machine1DataTable.Columns.Add("Thời gian", typeof(string));
            }

            // Thêm các cột từ thanh ghi vào DataTable trước cột thời gian
            foreach (int address in registerAddresses)
            {
                string columnName = $"Register_{address}";
                // Kiểm tra nếu cột chưa tồn tại thì mới thêm mới
                if (!machine1DataTable.Columns.Contains(columnName))
                {
                    machine1DataTable.Columns.Add(columnName, typeof(float));
                }
            }

            DataRow row = machine1DataTable.NewRow();

            // Ghi lại thời gian thực vào cột Thời gian
            row["Thời gian"] = DateTime.Now.ToString("HH:mm:ss");

            // Thêm dữ liệu từ các thanh ghi vào các cột tương ứng
            for (int i = 0; i < data.Length; i++)
            {
                string columnName = $"Register_{registerAddresses[i]}";
                float roundedValue;
                switch (registerAddresses[i])
                {
                    case 0:
                    case 6:
                        roundedValue = (float)Math.Round(data[i], 2);
                        break;
                    case 2:
                    case 8:
                        roundedValue = (float)Math.Round(data[i], 0);
                        break;
                    case 4:
                    case 10:
                        roundedValue = (float)Math.Round(data[i], 1);
                        break;
                    default:
                        roundedValue = data[i];
                        break;
                }
                row[columnName] = roundedValue;
            }

            machine1DataTable.Rows.Add(row);

            dgvDataIp1.DataSource = machine1DataTable;

            FormatDataGridViewCabinet1(dgvDataIp1);

            dgvDataIp1.FirstDisplayedScrollingRowIndex = dgvDataIp1.Rows.Count - 1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="registerAddresses"></param>
        private void UpdateIp2DataTable(float[] data, int[] registerAddresses)
        {
            // Ensure "Thời gian" column exists
            if (!machine2DataTable.Columns.Contains("Thời gian"))
                machine2DataTable.Columns.Add("Thời gian", typeof(string));

            // Add columns from register addresses
            foreach (int address in registerAddresses)
            {
                string columnName = $"Register_{address}";
                if (!machine2DataTable.Columns.Contains(columnName))
                    machine2DataTable.Columns.Add(columnName, typeof(float));
            }

            DataRow row = machine2DataTable.NewRow();
            row["Thời gian"] = DateTime.Now.ToString("HH:mm:ss");

            for (int i = 0; i < data.Length; i++)
            {
                string columnName = $"Register_{registerAddresses[i]}";
                float roundedValue = RoundValueBasedOnAddress(data[i], registerAddresses[i]);
                row[columnName] = roundedValue;
            }

            machine2DataTable.Rows.Add(row);
            dgvDataIp2.DataSource = machine2DataTable;
            FormatDataGridViewCabinet2(dgvDataIp2);
            dgvDataIp2.FirstDisplayedScrollingRowIndex = dgvDataIp2.Rows.Count - 1;
        }

        private float RoundValueBasedOnAddress(float value, int address)
        {
            switch (address)
            {
                case 0:
                case 8:
                case 10:
                case 12:
                case 24:
                case 26:
                case 28:
                case 16:
                    return (float)Math.Round(value, 2);

                case 2:
                case 4:
                case 6:
                case 18:
                case 20:
                case 22:
                    return (float)Math.Round(value, 0);

                case 14:
                case 30:
                    return (float)Math.Round(value, 1);

                default:
                    return value;
            }
        }


        /// <summary>
        /// Biến theo dõi phương thức cập nhật định kì
        /// </summary>
        private bool isIp1PeriodicStarted = false;
        private bool isIp2PeriodicStarted = false;

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private async Task StartPeriodicIp1(int timeDelay) 
        {
            string ipAddress = cbxListIp1.SelectedItem.ToString();

            if (!isIp1PeriodicStarted)
            {
                isIp1PeriodicStarted = true;

                while (shouldContinueIp1)
                {
                    await Task.Delay(timeDelay * 1000);

                    await Task.Run(() => ReadAndUpdateIp1Data());
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private async Task StartPeriodicIp2(int timeDelay)
        {
            string ipAddress = cbxListIp2.SelectedItem.ToString();

            if (!isIp2PeriodicStarted)
            {
                isIp2PeriodicStarted = true;

                while (shouldContinueIp2)
                {
                    await Task.Delay(timeDelay * 1000);

                    await Task.Run(() => ReadAndUpdateIp2Data());
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ipAddress"></param>
        /// <returns></returns>
        private bool ConnectToDeltaDevice(string ipAddress)
        {
            try
            {
                if (modbusClient != null)
                {
                    modbusClient.Disconnect();
                    modbusClient = null;
                }

                modbusClient = new ModbusClient(ipAddress, 502)
                {
                    ConnectionTimeout = 15000,
                };

                modbusClient.Connect();
                return true;
            }

            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Ghép byte theo định dạng LittleEndian để tạo ra số float
        /// </summary>
        /// <param name="val1"></param>
        /// <param name="val2"></param>
        /// <returns></returns>
        private float Int16ConvertToFloatLittleEndian(Int16 val1, Int16 val2)
        {
            byte[] temp = new byte[4];

            byte[] bytesVal1 = BitConverter.GetBytes(val1);
            byte[] bytesVal2 = BitConverter.GetBytes(val2);

            temp[0] = bytesVal1[0];
            temp[1] = bytesVal1[1];
            temp[2] = bytesVal2[0];
            temp[3] = bytesVal2[1];

            return BitConverter.ToSingle(temp, 0);
        }

        /// <summary>
        /// Quét thiết bị Delta
        /// </summary>
        /// <param name="networkPrefixes"></param>
        /// <returns></returns>
        private async Task<List<string>> FindDeltaDevicesAsync(params string[] networkPrefixes)
        {
            DeviceScanner scanner = new DeviceScanner();
            List<string> deltaDevices = await scanner.ScanDeltaDevices(networkPrefixes);
            return deltaDevices;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="comboBox"></param>
        /// <param name="networkPrefixes"></param>
        /// <returns></returns>
        private async Task UpdateDeltaDevicesComboBoxAsync(System.Windows.Forms.ComboBox comboBox, params string[] networkPrefixes)
        {
            List<string> deltaDevices = await FindDeltaDevicesAsync(networkPrefixes);

            int deviceCount = deltaDevices.Count;

            comboBox.BeginInvoke(new Action(() =>
            {
                comboBox.Items.Clear();

                if (deviceCount > 0)
                {
                    foreach (string ip in deltaDevices)
                    {
                        if (!connectedDevices.ContainsKey(ip))
                        {
                            comboBox.Items.Add(ip);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Không tìm thấy thiết bị Delta", "Lỗi", MessageBoxButtons.OK);
                }
            }));
        }

        /// <summary>
        /// Đọc dữ liệu từ thanh ghi
        /// </summary>
        /// <param name="registerAddresses"></param>
        /// <returns></returns>
        private float[] ReadDataFromRegisters(int[] registerAddresses)
        {
            List<float> resultData = new List<float>();

            try
            {
                if (modbusClient != null && modbusClient.Connected)
                {
                    foreach (int address in registerAddresses)
                    {
                        int[] singleData = modbusClient.ReadHoldingRegisters(address, 2);

                        if (singleData.Length >= 2)
                        {
                            float floatValue = Int16ConvertToFloatLittleEndian((short)singleData[0], (short)singleData[1]);

                            resultData.Add(floatValue);
                        }
                        else
                        {
                            if (!isErrorMessageBoxShown)
                            {
                                isErrorMessageBoxShown = true; 

                                Thread messageThread = new Thread(() =>
                                {
                                    MessageBox.Show("Không thể đọc dữ liệu từ thanh ghi Modbus.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });

                                messageThread.Start();
                            }
                        }
                    }
                }
                else
                {
                    if (!isErrorMessageBoxShown)
                    {
                        isErrorMessageBoxShown = true;

                        Thread messageThread = new Thread(() =>
                        {
                            MessageBox.Show("Kết nối đã bị đóng hoặc có sự cố.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        messageThread.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                if (!isErrorMessageBoxShown)
                {
                    isErrorMessageBoxShown = true; 

                    Thread messageThread = new Thread(() =>
                    {
                        MessageBox.Show("Lỗi khi đọc dữ liệu từ thanh ghi Modbus: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                    messageThread.Start();
                }
            }

            return resultData.ToArray();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCalculatorEnergy_Click(object sender, EventArgs e)
        {
            SimpleButton btn = sender as SimpleButton;
            if (btn != null)
            {
                DataTable selectedDataTable = null;
                Label lblTotalEnergy = null;
                DateTime startTime;
                DateTime endTime;
                string totalEnergy2Register;

                if (btn == btnCalculatorEnergy1)
                {
                    selectedDataTable = machine1DataTable;
                    lblTotalEnergy = lblTotalEnergyIp1;
                    startTime = dtpStartTimeIp1.Value;
                    endTime = dtpEndTimeIp1.Value;
                    totalEnergy2Register = "Register_6";
                }
                else if (btn == btnCalculatorEnergy2)
                {
                    selectedDataTable = machine2DataTable;
                    lblTotalEnergy = lblTotalEnergyIp2;
                    startTime = dtpStartTimeIp2.Value;
                    endTime = dtpEndTimeIp2.Value;
                    totalEnergy2Register = "Register_16";
                }
                else
                {
                    return;
                }

                float totalEnergy1 = EnergyCalculator.CalculateTotalEnergy(selectedDataTable, startTime, endTime, "Register_0");
                float totalEnergy2 = EnergyCalculator.CalculateTotalEnergy(selectedDataTable, startTime, endTime, totalEnergy2Register);

                if (totalEnergy1 == -11)
                {
                    lblTotalEnergy.Text = "Dữ liệu trống";
                }
                else if (totalEnergy1 == -12 || totalEnergy2 == -12)
                {
                    lblTotalEnergy.Text = "Thời gian không hợp lệ";
                    lblTotalEnergy.ForeColor = Color.Red;
                }
                else
                {
                    lblTotalEnergy.Text = $"Tổng tiêu thụ: {totalEnergy1} Kwh - {totalEnergy2} Kwh";
                    lblTotalEnergy.ForeColor = Color.Black;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void lblDataIp1_TextChanged(object sender, EventArgs e)
        {
            await Task.Run(() => CheckAndCalculateHourlyEnergy1());
        }

        private async void lblDataIp2_TextChanged(object sender, EventArgs e)
        {
            await Task.Run(() => CheckAndCalculateHourlyEnergy2());
        }

        /// <summary>
        /// 
        /// </summary>
        private string previousSelectedItem = null;
        private void cbxListIp1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxListIp1.SelectedItem != null)
            {
                string selectedItem = cbxListIp1.SelectedItem.ToString();

                if (selectedItem != previousSelectedItem)
                {
                    lblDataIp1.Text = "Cabinet 1";
                    lblDataIp1.ForeColor = Color.Black;

                    // Kiểm tra nếu cabinet1DataTable đã có dữ liệu, thì xóa hết
                    if (cabinet1DataTable.Rows.Count > 0)
                    {
                        cabinet1DataTable.Rows.Clear();
                        MessageBox.Show("Dữ liệu trong cabinet1DataTable đã được xóa.");
                    }

                    dgvDataIp1.DataSource = cabinet1DataTable;
                    previousSelectedItem = selectedItem;
                }
            }
        }

        private void xtraTabControl1_Selected(object sender, DevExpress.XtraTab.TabPageEventArgs e)
        {

        }

        private void lblStatusIp2_TextChanged(object sender, EventArgs e)
        {
            if(lblStatusIp2.Text == "Kết nối thành công")
            {
                // Lấy thời gian hiện tại với giây được thiết lập thành 00
                DateTime currentTime = DateTime.Now;
                DateTime roundedTime = new DateTime(currentTime.Year, currentTime.Month, currentTime.Day, currentTime.Hour, currentTime.Minute, 0);

                // Đặt giá trị cho DateTimePicker
                dtpStartTimeIp2.Value = roundedTime;
            }
        }

        private void lblStatusIp1_TextChanged(object sender, EventArgs e)
        {
            if (lblStatusIp1.Text == "Kết nối thành công")
            {
                // Lấy thời gian hiện tại với giây được thiết lập thành 00
                DateTime currentTime = DateTime.Now;
                DateTime roundedTime = new DateTime(currentTime.Year, currentTime.Month, currentTime.Day, currentTime.Hour, currentTime.Minute, 0);

                // Đặt giá trị cho DateTimePicker
                dtpStartTimeIp1.Value = roundedTime;
            }
        }
    }
}
