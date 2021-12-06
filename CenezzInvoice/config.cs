using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenezzInvoice
{
    class config
    {

        public static TextReader tr = new StreamReader("" + Path.GetDirectoryName(Application.ExecutablePath) + '/' + "config.ini");
        static string x = tr.ReadToEnd();

        public static string[] vector = x.Split(new char[] { '\r' });
        public static string srv = "" + vector[0].Replace("\r", "").Replace("\n", "");
        public static string usr = "" + vector[1].Replace("\r", "").Replace("\n", "");
        public static string pss = "" + vector[2].Replace("\r", "").Replace("\n", "");
        public static string dbb = "" + vector[3].Replace("\r", "").Replace("\n", "");
        public static string porto = "" + vector[4].Replace("\r", "").Replace("\n", "");
        public static string numemp = "" + vector[5].Replace("\r", "").Replace("\n", "");
        public static string prefix = "" + vector[6].Replace("\r", "").Replace("\n", "");
        public static string logeded = "0";
        public static string idinvoice = "";
        public static string lvl = "0";
        public static string almacenado = "";
        public static string tempofiles = "" + System.IO.Path.GetTempPath();

        public static string cade = @"Server=" + srv + "," + porto + ";Database=" + dbb + ";User Id=" + usr + ";Password=" + pss + ";MultipleActiveResultSets=true;";
        public static SqlConnection conn = new SqlConnection(@"" + cade);

        public static string bcobra = "0";
        public static string bmul = "0";
        public static string bcompa = "0";
        public static string btora = "0";
        public static string bextracto = "0";
        public static string bxsurtir = "0";
        public static string bventas = "0";
        public static string bconex = "0";
        public static string vcosto = "0";
        public static string vprecio = "0";
        public static string mcosto = "0";



        public static bool GetDateFormat(DateTime startDate, DateTime endDate, out string mensaje)
        {

            bool status = false;
            /*Valida fecha*/
            if (startDate.Date <= endDate.Date)
            {
                status = true;
                TimeSpan difference = endDate.Subtract(startDate.Date);

                StringBuilder sb = new StringBuilder();

                if (difference.Ticks == 0)
                {
                    sb.Append("1");
                }
                else if (difference.Ticks > 0)
                {
                    // This is to convert the timespan to datetime object
                    DateTime totalDate = DateTime.MinValue + difference;

                    int differenceInYears = totalDate.Year - 1;
                    int differenceInMonths = totalDate.Month;
                    int differenceInDays = totalDate.Day - 1;

                    if (differenceInYears > 0)
                        //sb.AppendFormat("{0} año(s)", differenceInYears);
                        differenceInMonths = differenceInMonths + (differenceInYears * 12);
                    if (differenceInMonths > 0)
                        if (differenceInMonths == 1)
                            sb.AppendFormat("{0}", differenceInMonths);
                        else
                            sb.AppendFormat("{0}", differenceInMonths);
                    /*
                    if (differenceInDays > 0)
                        if (differenceInDays == 1)
                            sb.AppendFormat(" {0} día", differenceInDays);
                        else
                            sb.AppendFormat(" {0} días", differenceInDays);
                    */
                }

                mensaje = sb.ToString();
            }
            else
            {
                mensaje = "Error";
            }
            return status;

        }
        public static int MonthDiff(DateTime startDate, DateTime endDate)
        {
            int months = 0;
            if (startDate > endDate)
            {
                months = -1;
            }
            else
            {
                months = ((endDate.Year * 12) + endDate.Month) - ((startDate.Year * 12) + startDate.Month);
                // if (endDate.Day >= startDate.Day)
                // {
                months = months + 1;
                // }
            }
            return months;
        }

        public static void BuildHeader()
        {

        }
    }
}
