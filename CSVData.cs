using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;


namespace DeliveryDashboardCSV
{
    public class CSVData
    {
        SqlConnection SQLcon = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["DeliveryDashboard"].ConnectionString);
        DateTime dtFinYearStart;
        DateTime dtFinYearEnd;
        int intMonths;
        string strForecastMonths;
        string strActualMonths;
        List<string> months = new List<string>();
        List<string> actualMonths = new List<string>();
        DataTable dt = new DataTable();

        protected void GetFinancialYear()
        {
            DateTime dtCurrentYear = DateTime.Now;
            DateTime dtPreviousYear = dtCurrentYear.AddYears(-1);
            DateTime dtNextYear = dtCurrentYear.AddYears(1);

            string strFinYearStart;
            string strFinYearEnd;

            if (dtCurrentYear.Month > 3)
            {
                strFinYearStart = "04/" + dtCurrentYear.Year.ToString();
                dtFinYearStart = Convert.ToDateTime(strFinYearStart);

                strFinYearEnd = "04/" + dtNextYear.Year.ToString();
                dtFinYearEnd = Convert.ToDateTime(strFinYearEnd);
            }
            else
            {
                strFinYearStart = "04/" + dtPreviousYear.Year.ToString();
                dtFinYearStart = Convert.ToDateTime(strFinYearStart);

                strFinYearEnd = "04/" + dtCurrentYear.Year.ToString();
                dtFinYearEnd = Convert.ToDateTime(strFinYearEnd);
            }
        }

        protected void GetNoMonths(DateTime Start, DateTime End)
        {
            int iNoMonths = ((End.Year - Start.Year) * 12) + End.Month - Start.Month;
            iNoMonths++;
            intMonths = iNoMonths;
        }

        protected void GetForecastMonths()
        {
            DateTime dtCurrentDate = Convert.ToDateTime(DateTime.Now.ToString("MMM-yy"));
            DateTime date;
            strForecastMonths = "";

            //Months that use Forecasts
            GetNoMonths(dtCurrentDate, dtFinYearEnd);
            date = dtCurrentDate;

            for (int i = 0; i < intMonths; i++)
            {
                strForecastMonths += ", SM.[" + date.ToString("MMM-yy").ToUpper() + "]";
                months.Add(", SM.[" + date.ToString("MMM-yy").ToUpper() + "]");
                date = date.AddMonths(1);
            }
        }

        protected void GetActualMonths()
        {
            DateTime dtCurrentDate = Convert.ToDateTime(DateTime.Now.ToString("MMM-yy"));
            DateTime date;
            strActualMonths = "";
            actualMonths.Clear();

            //Months that use Actuals
            GetNoMonths(dtFinYearStart, dtCurrentDate);
            date = dtFinYearStart;
            intMonths = intMonths - 1;

            for (int i = 0; i < intMonths; i++)
            {
                strActualMonths += ", TM.[" + date.ToString("MMM-yy").ToUpper() + "]";
                actualMonths.Add(date.ToString("MMM-yy").ToUpper());
                date = date.AddMonths(1);
            }
        }

        protected void ActualIfPresent()
        {
            int i = 0;
            int intResourceID = 0;
            int intProjectID = 0;
            string strCommandString = "";

            foreach (DataRow row in dt.Rows)
            {
                foreach (string actualmonth in actualMonths)
                {
                    if ((row[actualmonth].ToString() == "0") || (row[actualmonth].ToString() == "&nbsp;"))
                    {
                        intResourceID = Convert.ToInt32(row["Resource_ID"].ToString());
                        intProjectID = Convert.ToInt32(row["Project_ID"].ToString());
                        strCommandString = "SELECT MonthlyTotals.[" + actualmonth + "] FROM MonthlyTotals, TotalLookUp WHERE MonthlyTotals.Total_ID = TotalLookUp.Total_ID AND TotalLookUp.IsActual = 'false' AND TotalLookUp.Project_ID = @projectid AND TotalLookUp.Resource_ID = @resourceid";
                        SqlCommand cmd = new SqlCommand(strCommandString, SQLcon);
                        cmd.Parameters.AddWithValue("@projectid", intProjectID);
                        cmd.Parameters.AddWithValue("@resourceid", intResourceID);
                        SQLcon.Open();
                        SqlDataReader reader;
                        reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            row[actualmonth] = reader[actualmonth];
                        }
                        SQLcon.Close();
                    }
                }
                i++;
            }


        }

        public string GetResourceCSV()
        {

            GetFinancialYear();
            GetActualMonths();
            GetForecastMonths();

            SqlCommand cmd = new SqlCommand("SELECT DISTINCT TM.Resource_ID, TM.Project_ID, TM.Client_Name, TM.Sector_Name, TM.Project_Name, TM.Role_Name, TM.Resource_Name, TM.Type_Name,  TM.Source_Name, TM.Resource_Buy, TM.Resource_Sell" + strActualMonths + strForecastMonths + " FROM (SELECT TL.Resource_ID, PI.Project_ID, CL.Client_Name, RC.Resource_Name, PR.Resource_Buy, PR.Resource_Sell, ER.Role_Name, ES.Source_Name, PI.Project_Name, PS.Sector_Name, PT.Type_Name" + strActualMonths + " FROM  ProjectInformation as PI, RateCard as RC, TotalLookup as TL, ProjectResources as PR, ProjectSector as PS, ProjectType as PT, EmployeeRole as ER, EmployeeSource as ES, MonthlyTotals as TM, Clients as CL WHERE  PI.Project_ID = TL.Project_ID AND PI.Project_Client = CL.Client_ID AND  TL.Resource_ID = RC.Resource_ID  AND  PR.Project_ID = TL.Project_ID  AND  PR.Resource_ID = RC.Resource_ID  AND  PI.Project_Sector = PS.Sector_ID AND  PI.Project_Type = PT.Type_ID  AND  TL.IsActual = 'true'  AND  TM.Total_ID = TL.Total_ID  AND RC.Resource_Type = 'Employee' AND ER.Role_ID = PR.Employee_Role AND ES.Source_ID = PR.Employee_Source) as TM  LEFT OUTER JOIN (SELECT  PI.Project_ID, RC.Resource_Name" + strForecastMonths + "  FROM  ProjectInformation as PI, RateCard as RC, TotalLookup as TL, ProjectResources as PR, ProjectSector as PS, ProjectType as PT, MonthlyTotals as SM  WHERE  PI.Project_ID = TL.Project_ID AND  TL.Resource_ID = RC.Resource_ID AND  PR.Project_ID = PI.Project_ID AND  PR.Project_ID = TL.Project_ID AND  PR.Resource_ID = RC.Resource_ID AND  PI.Project_Sector = PS.Sector_ID AND  PI.Project_Type = PT.Type_ID AND  TL.IsActual = 'false' AND  SM.Total_ID = TL.Total_ID AND RC.Resource_Type = 'Employee') as SM on TM.Resource_Name = SM.Resource_Name AND TM.Project_ID = SM.Project_ID ORDER BY TM.Client_Name, TM.Project_Name, TM.Resource_Name", SQLcon); SQLcon.Open();

            dt = new DataTable();
            SqlDataAdapter sqlAd = new SqlDataAdapter(cmd);
            sqlAd.Fill(dt);
            SQLcon.Close();

            ActualIfPresent();

            //Convert nulls to zero's
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    if (row.IsNull(col))
                        row.SetField(col, 0);
                }
            }

            //Remove unnecessary columns
            dt.Columns.Remove("Resource_ID");
            dt.Columns.Remove("Project_ID");

            dt.Columns.Add("Margin", typeof(double), "((Resource_Sell-Resource_Buy)/Resource_Sell)*100");
            dt.Columns.Add("Margin2");
            dt.Columns.Add("Resource_Sell_Euro", typeof(double), "Resource_Sell*1.26");
            dt.Columns.Add("Blank_Column");
            dt.Columns.Add("Total_Days");
            dt.Columns.Add("Total_Days2");
            dt.SetColumnsOrder1("Client_Name", "Sector_Name", "Project_Name", "Type_Name", "Role_Name", "Resource_Name", "Source_Name", "Resource_Sell_Euro", "Resource_Sell", "Blank_Column", "Resource_Buy", "Margin2", "Total_Days", "Total_Days2");

            int intRowIndex = 0;

            foreach (DataRow row in dt.Rows)
            {
                double dblRowTotal = 0;
                for (int i = 14; i <= 26; i++)
                {
                    dblRowTotal += (double)(dt.Rows[intRowIndex][i]);
                }
                dt.Rows[intRowIndex]["Total_Days"] = dblRowTotal;
                dt.Rows[intRowIndex]["Total_Days2"] = dblRowTotal;
                dt.Rows[intRowIndex]["Margin2"] = (dt.Rows[intRowIndex]["Margin"]);
                if ((string)dt.Rows[intRowIndex]["Margin2"] == "∞" || (string)dt.Rows[intRowIndex]["Margin2"] == "-∞" || (string)dt.Rows[intRowIndex]["Margin2"] == "NaN" || (string)dt.Rows[intRowIndex]["Margin2"] == "Infinity" || (string)dt.Rows[intRowIndex]["Margin2"] == "-Infinity")
                {
                    dt.Rows[intRowIndex]["Margin2"] = "0";
                }
                else
                {
                    dt.Rows[intRowIndex]["Margin2"] = Math.Round(Convert.ToDouble(dt.Rows[intRowIndex]["Margin2"].ToString()), 0);
                }
                intRowIndex++;
            }

            dt.Columns.Remove("Margin");

            //Rename columns
            dt.Columns["Resource_Name"].ColumnName = "Name";
            dt.Columns["Project_Name"].ColumnName = "Project Name";
            dt.Columns["Sector_Name"].ColumnName = "Sector";
            dt.Columns["Client_Name"].ColumnName = "Client";
            dt.Columns["Type_Name"].ColumnName = "Type";
            dt.Columns["Role_Name"].ColumnName = "Role";
            dt.Columns["Source_Name"].ColumnName = "Source";
            dt.Columns["Resource_Buy"].ColumnName = "Pay (GBP)";
            dt.Columns["Resource_Sell"].ColumnName = "Sell (GBP)";
            dt.Columns["Resource_Sell_Euro"].ColumnName = "Sell (EUR)";
            dt.Columns["Blank_Column"].ColumnName = " ";
            dt.Columns["Total_Days"].ColumnName = "Total Days";
            dt.Columns["Total_Days2"].ColumnName = "Currently Maped Out";
            dt.Columns["Margin2"].ColumnName = "Margin (%)";

            StringBuilder sBuilder = new System.Text.StringBuilder();
            IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
            sBuilder.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sBuilder.AppendLine(string.Join(",", fields));
            }

            return sBuilder.ToString();
        }
    }

    public static class DataTableExtensions
    {
        public static void SetColumnsOrder1(this DataTable table, params String[] columnNames)
        {
            int columnIndex = 0;
            foreach (var columnName in columnNames)
            {
                table.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }
        }
    }
}