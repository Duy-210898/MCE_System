using System;
using System.Data;
using System.Windows.Forms;


public static class ControlExtensions
{
    public static void DoubleBuffered(this Control control, bool enable)
    {
        var doubleBufferPropertyInfo = control.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        doubleBufferPropertyInfo?.SetValue(control, enable, null);
    }
}

public class HourlyConsumptionCalculator
{
    private DataTable hourlyCabinet1DataTable;
    private DataTable machine1DataTable;
    public DataTable HourlyCabinet1DataTable => hourlyCabinet1DataTable;

    public HourlyConsumptionCalculator(DataTable machine1Data)
    {
        // Tạo DataTable cho cả hai máy
        hourlyCabinet1DataTable = new DataTable();
        hourlyCabinet1DataTable.Columns.Add("Thời gian", typeof(string));
        hourlyCabinet1DataTable.Columns.Add("Tiêu thụ máy 1", typeof(float));
        hourlyCabinet1DataTable.Columns.Add("Tiêu thụ máy 2", typeof(float));

        machine1DataTable = machine1Data;
    }

    public void StartHourlyProcessing()
    {
        ScheduleHourlyTask(DateTime.Now.AddMinutes(30), this);
    }

    private void ScheduleHourlyTask(DateTime scheduledTime, HourlyConsumptionCalculator calculator)
    {
        TimeSpan delay = scheduledTime - DateTime.Now;
        if (delay.TotalMilliseconds < 0)
        {
            return;
        }

        System.Timers.Timer timer = new System.Timers.Timer();
        timer.Interval = delay.TotalMilliseconds;
        timer.Elapsed += (sender, e) => calculator.CalculateAndFillHourlyData(scheduledTime, machine1DataTable);
        timer.Start();
    }

    private void CalculateAndFillHourlyData(DateTime scheduledTime, DataTable machine1Data) 
    {
        DateTime startTime = scheduledTime.AddHours(-1).AddMinutes(30);
        DateTime endTime = scheduledTime;

        string timeRange = $"{startTime.ToString("HH:mm")} - {endTime.ToString("HH:mm")}";

        float machine1Consumption = FindValueAtTime(machine1Data, startTime, 0) - FindValueAtTime(machine1Data, endTime, 0);
        hourlyCabinet1DataTable.Rows.Add(timeRange, machine1Consumption);
    }

    private float FindValueAtTime(DataTable dataTable, DateTime time, int register)
    {
        foreach (DataRow row in dataTable.Rows)
        {
            DateTime rowTime;
            if (!DateTime.TryParseExact(row["Thời gian"].ToString(), "HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out rowTime))
            {
                continue;
            }

            if (rowTime.Hour == time.Hour && rowTime.Minute == time.Minute)
            {
                float value;
                // Lấy giá trị từ cột Register_0 hoặc Register_6 tùy thuộc vào tham số register
                if (!float.TryParse(row[$"Register_{register}"].ToString(), out value))
                {
                    return 0;
                }
                return value;
            }
        }
        return 0;
    }
}
