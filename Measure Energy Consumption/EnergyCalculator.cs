using System;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Measure_Energy_Consumption
{
    public class EnergyCalculator
    {
        /// <summary>
        /// Hàm tính điện năng tiêu thụ trong khoảng thời gian được chọn
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static float CalculateTotalEnergy(DataTable dataTable, DateTime startTime, DateTime endTime, string columnName)
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                return -11;
            }

            DataRow startRow = dataTable.AsEnumerable()
                                        .FirstOrDefault(row => DateTime.TryParse(row["Thời gian"].ToString(), out DateTime timeStamp)
                                                               && timeStamp.TimeOfDay.Hours == startTime.Hour
                                                               && timeStamp.TimeOfDay.Minutes == startTime.Minute);

            DataRow endRow = dataTable.AsEnumerable()
                                      .FirstOrDefault(row => DateTime.TryParse(row["Thời gian"].ToString(), out DateTime timeStamp)
                                                             && timeStamp.TimeOfDay.Hours == endTime.Hour
                                                             && timeStamp.TimeOfDay.Minutes == endTime.Minute);

            if (startRow != null && endRow != null)
            {
                float startValue = Convert.ToSingle(startRow[columnName]);
                float endValue = Convert.ToSingle(endRow[columnName]);
                float totalEnergy = Convert.ToSingle(Math.Round(endValue - startValue, 2));
                return totalEnergy;
            }

            return -12;
        }
    }
}
