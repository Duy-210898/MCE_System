using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using System.Threading;

namespace Measure_Energy_Consumption
{
    public class ExcelUpdate
    {
        private static bool isMessageBoxShown = false;

        /// <summary>
        /// Tìm file trong đường dẫn
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <param name="fileNamePattern"></param>
        /// <returns></returns>
        public static string FindExcelFile(string directoryPath, string fileNamePattern)
        {
            string[] excelFiles = Directory.GetFiles(directoryPath, fileNamePattern);
            return excelFiles.FirstOrDefault();
        }


        /// <summary>
        /// Biến kiểm tra file excel đang mở từ excel
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
        }

        /// <summary>
        /// Đảm bảo đường dẫn tồn tại, nếu chưa thì tạo folder mới
        /// </summary>
        /// <param name="directoryPath"></param>
        public static void EnsureDirectoryExists(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                try
                {
                    Directory.CreateDirectory(directoryPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error creating directory: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Cập nhật dữ liệu điện năng tiêu thụ theo giờ vào sheet Hourly
        /// </summary>
        /// <param name="totalEnergy1"></param>
        /// <param name="totalEnergy2"></param>
        /// <param name="startTime"></param>
        /// <param name="currentTime"></param>
        /// <param name="cabinet"></param>
        public static void UpdateExcelWithHourlyData(float totalEnergy1, float totalEnergy2, DateTime startTime, DateTime currentTime, string cabinet)
        {
            string mceFolderPath = @"D:\MCE Data";
            string dailyFolderPath = Path.Combine(mceFolderPath, "Daily");
            string fileNamePattern = $"{cabinet} - {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx";

            EnsureDirectoryExists(dailyFolderPath);

            string filePath = FindExcelFile(dailyFolderPath, fileNamePattern);
            //  Trường hợp không tìm thấy đường dẫn
            if (string.IsNullOrEmpty(filePath))
            {
                filePath = Path.Combine(dailyFolderPath, fileNamePattern);

                // Tạo một file Excel mới
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Hourly");

                    // Đặt tên cho các cột
                    List<string> columnNames = new List<string>
                    {
                        "Thời gian",
                        "Total Energy 1\n(Kwh)",
                        "Total Energy 2\n(Kwh)"
                    };

                    // Đặt tiêu đề cho cột
                    for (int i = 0; i < columnNames.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columnNames[i];
                    }

                    // Tạo chuỗi thời gian cho dòng hiện tại
                    string timeRange = $"{startTime.ToString("HH:mm")} - {currentTime.ToString("HH:mm")}";

                    // Thêm dữ liệu vào sheet Hourly
                    worksheet.Cells[2, 1].Value = timeRange;
                    worksheet.Cells[2, 2].Value = totalEnergy1;
                    worksheet.Cells[2, 3].Value = totalEnergy2;

                    // Áp dụng định dạng từ ExcelFormatter
                    ExcelFormatter.FormatStyle(worksheet);
                    ExcelFormatter.RoundExcelHourlyColumns(worksheet);

                    // Lưu file Excel
                    FileInfo excelFile = new FileInfo(filePath);
                    excelPackage.SaveAs(excelFile);
                }
            }
            else
            {
                // Nếu file không mở từ ứng dụng khác
                if (!IsFileLocked(filePath))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "Hourly");

                        if (worksheet == null)
                        {
                            // Tạo một worksheet mới nếu không tìm thấy
                            worksheet = excelPackage.Workbook.Worksheets.Add("Hourly");

                            // Đặt tên cho các cột
                            List<string> columnNames = new List<string>
                            {
                                "Thời gian",
                                "Total Energy 1\n(Kwh)",
                                "Total Energy 2\n(Kwh)"
                            };

                            // Đặt tiêu đề cho cột
                            for (int i = 0; i < columnNames.Count; i++)
                            {
                                worksheet.Cells[1, i + 1].Value = columnNames[i];
                            }
                        }

                        // Tìm dòng cuối cùng đã có dữ liệu
                        int lastRow = worksheet.Dimension?.End.Row + 1 ?? 2;

                        // Tạo chuỗi thời gian cho dòng hiện tại
                        string timeRange = $"{startTime.ToString("HH:mm")} - {currentTime.ToString("HH:mm")}";

                        // Thêm dữ liệu vào sheet Hourly
                        worksheet.Cells[lastRow, 1].Value = timeRange;
                        worksheet.Cells[lastRow, 2].Value = totalEnergy1;
                        worksheet.Cells[lastRow, 3].Value = totalEnergy2;

                        // Áp dụng định dạng từ ExcelFormatter
                        ExcelFormatter.FormatStyle(worksheet);
                        // Làm tròn số
                        ExcelFormatter.RoundExcelHourlyColumns(worksheet);

                        excelPackage.Save();
                    }
                }
                else
                {
                    if (!isMessageBoxShown)
                    {
                        Thread messageThread = new Thread(() =>
                        {
                            DialogResult dialogResult = MessageBox.Show("Không thể cập nhật dữ liệu vào file Excel vì file đang được sử dụng bởi một ứng dụng khác.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            if (dialogResult == DialogResult.OK)
                            {
                                isMessageBoxShown = false;
                            }
                            else
                            {
                                isMessageBoxShown = true;
                            }
                        });
                        messageThread.Start();

                        isMessageBoxShown = true;
                    }
                }
            }
        }

        /// <summary>
        /// Cập nhật dữ liệu realtime từ thanh ghi vào sheet Cabinet
        /// </summary>
        /// <param name="newData"></param>
        /// <param name="cabinet"></param>
        public static void UpdateExcelWithNewData(float[] newData, string cabinet)
        {
            string mceFolderPath = @"D:\MCE Data";
            string dailyFolderPath = Path.Combine(mceFolderPath, "Daily");
            string fileNamePattern = $"{cabinet} - {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx";

            EnsureDirectoryExists(dailyFolderPath);

            string filePath = FindExcelFile(dailyFolderPath, fileNamePattern);

            if (string.IsNullOrEmpty(filePath))
            {
                filePath = Path.Combine(dailyFolderPath, fileNamePattern);

                // Tạo một file Excel mới
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(cabinet);

                    // Đặt tên cho các cột
                    List<string> columnNames = new List<string>
                    {
                        "Thời gian",
                        "Điện năng\n(Kwh)",
                        "Điện áp\n(V)",
                        "Dòng điện\n(A)",
                        "Điện năng\n(Kwh)",
                        "Điện áp\n(V)",
                        "Dòng điện\n(A)"
                    };

                    // Đặt tiêu đề cho cột
                    for (int i = 0; i < columnNames.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columnNames[i];
                    }

                    // Đặt dữ liệu mới vào file Excel
                    worksheet.Cells[2, 1].Value = DateTime.Now.ToString("HH:mm:ss");
                    for (int i = 0; i < newData.Length; i++)
                    {
                        worksheet.Cells[2, i + 2].Value = newData[i];
                    }

                    // Áp dụng định dạng từ ExcelFormatter
                    ExcelFormatter.FormatStyle(worksheet);
                    ExcelFormatter.RoundExcelColumns(worksheet);

                    // Lưu file Excel
                    FileInfo excelFile = new FileInfo(filePath);
                    excelPackage.SaveAs(excelFile);
                }
            }
            else
            {
                if (!IsFileLocked(filePath))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == cabinet);

                        if (worksheet == null)
                        {
                            // Tạo một worksheet mới nếu không tìm thấy
                            worksheet = excelPackage.Workbook.Worksheets.Add(cabinet);

                            // Đặt tên cho các cột
                            List<string> columnNames = new List<string>
                            {
                                "Thời gian",
                                "Điện năng\n(Kwh)",
                                "Điện áp\n(V)",
                                "Dòng điện\n(A)",
                                "Điện năng\n(Kwh)",
                                "Điện áp\n(V)",
                                "Dòng điện\n(A)"
                            };

                            // Đặt tiêu đề cho cột
                            for (int i = 0; i < columnNames.Count; i++)
                            {
                                worksheet.Cells[1, i + 1].Value = columnNames[i];
                            }
                        }

                        // Tìm dòng cuối cùng đã có dữ liệu
                        int lastRow = worksheet.Dimension?.End.Row + 1 ?? 2;

                        // Đặt dữ liệu mới vào file Excel
                        worksheet.Cells[lastRow, 1].Value = DateTime.Now.ToString("HH:mm:ss");
                        for (int i = 0; i < newData.Length; i++)
                        {
                            worksheet.Cells[lastRow, i + 2].Value = newData[i];
                        }

                        // Áp dụng định dạng từ ExcelFormatter
                        ExcelFormatter.FormatStyle(worksheet);
                        ExcelFormatter.RoundExcelColumns(worksheet);

                        excelPackage.Save();
                    }
                }
                else
                {
                    if (!isMessageBoxShown)
                    {
                        Thread messageThread = new Thread(() =>
                        {
                            DialogResult dialogResult = MessageBox.Show("Không thể cập nhật dữ liệu vào file Excel vì file đang được sử dụng bởi một ứng dụng khác.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            if (dialogResult == DialogResult.OK)
                            {
                                isMessageBoxShown = false;
                            }
                            else
                            {
                                isMessageBoxShown = true;
                            }
                        });
                        messageThread.Start();

                        isMessageBoxShown = true;
                    }
                }
            }
        }
    }

    public class ExcelFormatter
    {

        /// <summary>
        /// Làm tròn số trong sheet Cabinet
        /// </summary>
        /// <param name="worksheet"></param>
        public static void RoundExcelColumns(ExcelWorksheet worksheet)
        {
            // Cột 2 và 5: làm tròn về 2 chữ số phần thập phân
            worksheet.Column(2).Style.Numberformat.Format = "0.00";
            worksheet.Column(5).Style.Numberformat.Format = "0.00";

            // Cột 3 và 7: làm tròn về 1 chữ số phần thập phân
            worksheet.Column(3).Style.Numberformat.Format = "0";
            worksheet.Column(6).Style.Numberformat.Format = "0";

            // Cột 4 và 8: không làm tròn phần thập phân
            worksheet.Column(4).Style.Numberformat.Format = "0.0";
            worksheet.Column(7).Style.Numberformat.Format = "0.0";
        }

        /// <summary>
        /// Làm tròn số trong sheet Hourly
        /// </summary>
        /// <param name="worksheet"></param>
        public static void RoundExcelHourlyColumns(ExcelWorksheet worksheet)
        {
            // Cột 2 và 5: làm tròn về 2 chữ số phần thập phân
            worksheet.Column(2).Style.Numberformat.Format = "0.00";
            worksheet.Column(3).Style.Numberformat.Format = "0.00";
        }

        /// <summary>
        /// Định dạng file excel cơ bản
        /// </summary>
        /// <param name="worksheet"></param>
        public static void FormatStyle(ExcelWorksheet worksheet)
        {
            // Bật wrap text
            worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.WrapText = true;

            // Tắt gridline
            worksheet.View.ShowGridLines = false;

            // Căn giữa toàn bộ dữ liệu
            worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Kẻ viền cho các ô
            var border = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border;
            border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;

            // Tô màu xen kẽ từ dòng thứ 3
            for (int i = 3; i <= worksheet.Dimension.End.Row; i += 2)
            {
                using (var range = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                }
            }
            const int minimumColumnWidth = 15;
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                worksheet.Column(col).Width = Math.Max(minimumColumnWidth, worksheet.Column(col).Width);
            }
        }
    }
}
