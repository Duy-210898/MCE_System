using System;
using System.Threading;
using System.Windows.Forms;

namespace Measure_Energy_Consumption
{
    internal static class Program
    {
        private static Mutex mutex; // Sử dụng Mutex để đảm bảo chỉ chạy một instance của ứng dụng

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Kiểm tra xem ứng dụng đã được khởi chạy chưa
            bool createdNew;
            mutex = new Mutex(true, "Measure_Energy_Consumption", out createdNew);
            if (!createdNew)
            {
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new frmMain());
        }
    }
}
