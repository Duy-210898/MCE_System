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
using System.Windows.Ink;

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
                    ExcelFormatter.FormatStyle(worksheet, "Cabinet 1");
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
                        ExcelFormatter.FormatStyle(worksheet, cabinet);
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
        /// 
        private static void CreateNewExcelFile(string filePath, string cabinet, float[] newData, bool isCabinet2)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(cabinet);

                // Đặt tên cho các cột dựa vào giá trị của cabinet
                List<string> columnNames = GetColumnNamesForCabinet(cabinet);

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
                ExcelFormatter.FormatStyle(worksheet, cabinet);
                ExcelFormatter.RoundExcelColumns(worksheet, cabinet);

                // Lưu file Excel
                FileInfo excelFile = new FileInfo(filePath);
                excelPackage.SaveAs(excelFile);
            }
        }


        private static void HandleFileInUseError(string filePath)
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
        private static List<string> GetColumnNamesForCabinet(string cabinet)
        {
            List<string> columnNames = new List<string>();
            if (cabinet == "Cabinet 1")
            {
                columnNames.AddRange(new string[]
                {
            "Thời gian",
            "Điện năng\n(Kwh)",
            "Điện áp\n(V)",
            "Dòng điện\n(A)",
            "Điện năng\n(Kwh)",
            "Điện áp\n(V)",
            "Dòng điện\n(A)"
                });
            }
            else if (cabinet == "Cabinet 2")
            {
                columnNames.AddRange(new string[]
                {
            "Thời gian",
            "Tổng điện năng\nMáy 1\n(Kwh)",
            "Điện áp\nPha A\nMáy 1\n(V)",
            "Điện áp\nPha B\nMáy 1\n(V)",
            "Điện áp\nPha C\nMáy 1\n(V)",
            "Dòng điện\nPha A\nMáy 1\n(A)",
            "Dòng điện\nPha B\nMáy 1\n(A)",
            "Điện áp\nPha C\nMáy 1\n(A)",
            "Công suất\nMáy 1\n(Kw)",
            "Tổng điện năng\nMáy 2\n(Kwh)",
            "Điện áp\nPha A\nMáy 2\n(V)",
            "Điện áp\nPha B\nMáy 2\n(V)",
            "Điện áp\nPha C\nMáy 2\n(V)",
            "Dòng điện\nPha A\nMáy 2\n(A)",
            "Dòng điện\nPha B\nMáy 2\n(A)",
            "Điện áp\nPha C\nMáy 2\n(A)",
            "Công suất\nMáy 2\n(Kw)"
                });
            }
            return columnNames;
        }

        private static void UpdateExistingExcelFile(string filePath, string cabinet, float[] newData)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == cabinet);

                if (worksheet == null)
                {
                    // Tạo một worksheet mới nếu không tìm thấy
                    worksheet = excelPackage.Workbook.Worksheets.Add(cabinet);

                    // Đặt tên cho các cột dựa vào giá trị của cabinet
                    List<string> columnNames = GetColumnNamesForCabinet(cabinet);

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
                ExcelFormatter.FormatStyle(worksheet, "Cabinet 1");
                ExcelFormatter.RoundExcelColumns(worksheet, cabinet);

                excelPackage.Save();
            }
        }

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
                if (cabinet == "Cabinet 1")
                {
                    CreateNewExcelFile(filePath, cabinet, newData, false);
                }
                else if (cabinet == "Cabinet 2")
                {
                    CreateNewExcelFile(filePath, cabinet, newData, true);
                }
            }
            else
            {
                if (!IsFileLocked(filePath))
                {
                    UpdateExistingExcelFile(filePath, cabinet, newData);
                }
                else
                {
                    HandleFileInUseError(filePath);
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
        public static void RoundExcelColumns(ExcelWorksheet worksheet, string cabinet)
        {
            if(cabinet == "Cabinet 1")
            {
                worksheet.Column(2).Style.Numberformat.Format = "0.00";
                worksheet.Column(5).Style.Numberformat.Format = "0.00";

                // Cột 3 và 7: làm tròn về 1 chữ số phần thập phân
                worksheet.Column(3).Style.Numberformat.Format = "0";
                worksheet.Column(6).Style.Numberformat.Format = "0";

                // Cột 4 và 8: không làm tròn phần thập phân
                worksheet.Column(4).Style.Numberformat.Format = "0.0";
                worksheet.Column(7).Style.Numberformat.Format = "0.0";
            }

            else if(cabinet == "Cabinet 2")
            {
                worksheet.Column(2).Style.Numberformat.Format = "0.00";
                worksheet.Column(10).Style.Numberformat.Format = "0.00";

                // Cột 3 và 6: làm tròn về 1 chữ số phần thập phân
                worksheet.Column(3).Style.Numberformat.Format = "0";
                worksheet.Column(5).Style.Numberformat.Format = "0";
                worksheet.Column(5).Style.Numberformat.Format = "0";
                worksheet.Column(11).Style.Numberformat.Format = "0";
                worksheet.Column(12).Style.Numberformat.Format = "0";
                worksheet.Column(13).Style.Numberformat.Format = "0";

                // Cột 4 và 8: không làm tròn phần thập phân
                worksheet.Column(6).Style.Numberformat.Format = "0.0";
                worksheet.Column(7).Style.Numberformat.Format = "0.0";
                worksheet.Column(8).Style.Numberformat.Format = "0.0";
                worksheet.Column(9).Style.Numberformat.Format = "0.0";
                worksheet.Column(14).Style.Numberformat.Format = "0.0";
                worksheet.Column(15).Style.Numberformat.Format = "0.0";
                worksheet.Column(16).Style.Numberformat.Format = "0.0";
                worksheet.Column(17).Style.Numberformat.Format = "0.0";
            }
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
        public static void FormatStyle(ExcelWorksheet worksheet, string cabinet)
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

            if (cabinet == "Cabinet 1")
            {
                // Tô màu xen kẽ từ dòng thứ 3 cho tất cả các cột
                for (int i = 3; i <= worksheet.Dimension.End.Row; i += 2)
                {
                    using (var range = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                    }
                }
            }

            else if(cabinet == "Cabinet 2")
            {
                // Tô màu xen kẽ các dòng cho các cột từ 2 đến cột 9 màu AliceBlue và cột 10 đến cột 17 màu LightYellow
                for (int i = 3; i <= worksheet.Dimension.End.Row; i += 2)
                {
                    using (var range = worksheet.Cells[i, 1, i, 9])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                    }

                    using (var range = worksheet.Cells[i, 10, i, 17])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                    }
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
