using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Configuration;
using System.Diagnostics;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp;
using PdfSharp.Drawing.Layout;

namespace CenezzInvoice
{
    public partial class list : Form
    {
        public list()
        {
            InitializeComponent();
        }
        public static int deci =2;
        public class LayoutHelper
        {
            private readonly PdfDocument _document;
            private readonly XUnit _topPosition;
            private readonly XUnit _bottomMargin;
            private XUnit _currentPosition;

            public LayoutHelper(PdfDocument document, XUnit topPosition, XUnit bottomMargin)
            {
                _document = document;
                _topPosition = topPosition;
                _bottomMargin = bottomMargin;
                // Set a value outside the page - a new page will be created on the first request.
                _currentPosition = bottomMargin + 10000;
            }

            public XUnit GetLinePosition(XUnit requestedHeight)
            {
                return GetLinePosition(requestedHeight, -1f);
            }

            public XUnit GetLinePosition(XUnit requestedHeight, XUnit requiredHeight)
            {
                XUnit required = requiredHeight == -1f ? requestedHeight : requiredHeight;
                if (_currentPosition + required > _bottomMargin)
                    CreatePage();
                XUnit result = _currentPosition;
                _currentPosition += requestedHeight;
                return result;
            }

            public XGraphics Gfx { get; private set; }
            public PdfPage Page { get; private set; }

            void CreatePage()
            {
                Page = _document.AddPage();
                Page.Size = PageSize.Letter;
                Gfx = XGraphics.FromPdfPage(Page);
                _currentPosition = _topPosition;
            }
        }

        public void FillDataGridView()
        {
            BindingSource bSource = new BindingSource();
            bSource.DataSource = null;
            string inicial = "", final = "", cli = "", emi = "";
            SqlConnection con = new SqlConnection("" + config.cade);
            con.Open();
            inicial = "" + init.Text; final = "" + finit.Text;

            var primera = Convert.ToDateTime(inicial);
            var segunda = Convert.ToDateTime(final);
            inicial = "" + primera.ToString("yyyy-MM-dd");
            final = "" + segunda.ToString("yyyy-MM-dd");
            string range = "";

            cli = "" + numc.Text;
            emi = "" + nume.Text;

            if (cli != "") { cli = " AND  " + cli; }
            if (emi != "") { }
            range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE) ";
            SqlDataAdapter DA = new SqlDataAdapter();
            string sqlSelectAll = invoice_query + " " + range + ";";
            DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

            DataTable table = new DataTable();
            DA.Fill(table);
            bSource.DataSource = table;
            this.lister.DataSource = bSource;
            this.resizegrid();

            //here i checked number of rows of dt1 and it shows the correct value
        }

        internal static string invoice_query = "SELECT folio AS FOLIO, number AS NUM,empresa AS #Emp, idcli AS #Cli,nomcli AS Nombre, paiscli AS Pais, fecha as Fecha, albaran as Albaran, origdest as [Origen Destino],FORMAT( convert(numeric(18,5),replace(tot,',','')),'###,###,###.00000','ES-mx') AS Total, currency as Moneda ,id,stats AS [Status] FROM invoicespl";


        private void list_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void list_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (idinvo.Text != "")
            {
                logpesos.Text = "";
                #region init
                string id = "", numee = "", folio = "", empresa = "", idcli = "", nomcli = "", callecli = "", numcli = "", numclii = "", colcli = "", muncli = "", edocli = "", paiscli = "", fecha = "", albaran = "", origdest = "", tot = "", currency = "", cpcli = "",nifcli="";
                string ide = "", nome = "", nife = "", callee = "", nume = "", numie = "", cole = "", cde = "", estadoe = "", paise = "", cpe = "";
                string obs1 = "", obs2 = "", obs3 = "", obs4 = "", obs5 = "";
                string csymbol = "";
                double kgscontainer = 0, kgsacums = 0;
                string netofact = "", brutosfact="";

                double mtspartidas = 0, mtsacum = 0, kgs1caja = 0, mts1caja = 0, tonelaje = 0,tonelajecontainer=0;
                double tarimaskgs = 0, kgstar = 0, kilosXpallets = 0;


                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                DateTime hoy = DateTime.Now;
                string dates = hoy.Day.ToString() + "-" + hoy.Month.ToString() + "-" + hoy.Year.ToString() + "-" + hoy.Hour.ToString() + "-" + hoy.Minute.ToString() + "-" + hoy.Second.ToString();
                ExcelPackage excel = new ExcelPackage();
                excel.Workbook.Worksheets.Add("FACTURA");
                excel.Workbook.Worksheets.Add("PACKING LIST");
                //excel.Workbook.Worksheets.Add("PACKING LIST");
                //excel.Workbook.Worksheets.Add("PACKING LIST2");
                try
                {
                    Directory.EnumerateFiles(@"" + config.tempofiles, "INVOICE_PACKINGLIST_*.xlsx").ToList().ForEach(x => File.Delete(x));
                }
                catch { }
                FileInfo excelFile = new FileInfo(@"" + config.tempofiles + @"\INVOICE_PACKINGLIST_" + dates + ".xlsx");
                var proforma = excel.Workbook.Worksheets["FACTURA"];
                //var concentrado = excel.Workbook.Worksheets["FACTURA"];
                //var plist = excel.Workbook.Worksheets["PACKING LIST"];
                var plist2 = excel.Workbook.Worksheets["PACKING LIST"];
                double netofull = 0, brutofull = 0;
                string cellrow1 = "", cellrow2 = "";
                #endregion init


                #region headerproform
                /*
                concentrado.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                concentrado.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                */

                SqlConnection con = new SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT * FROM invoicespl WHERE id =" + idinvo.Text + ";";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {

                        id = "" + row["id"]; folio = "" + row["folio"]; empresa = "" + row["empresa"]; idcli = "" + row["idcli"];
                        nomcli = "" + row["nomcli"];
                        callecli = "" + row["callecli"];
                        numcli = "" + row["numcli"]; numclii = "" + row["numclii"]; colcli = "" + row["colcli"];
                        muncli = "" + row["muncli"];
                        edocli = "" + row["edocli"]; paiscli = "" + row["paiscli"];
                        fecha = "" + row["fecha"];
                        origdest = "" + row["origdest"]; tot = "" + row["tot"];
                        currency = "" + row["currency"];
                        obs1 = "" + row["obs1"];
                        obs2 = "" + row["obs2"];
                        obs3 = "" + row["obs3"];
                        obs4 = "" + row["obs4"];
                        obs5 = "" + row["obs5"];

                        albaran = "" + row["albaran"];
                        numee = "" + row["number"];

                        netofact = "" + row["pesoneto"]; brutosfact = "" + row["pesobruto"];
                    }
                }
                else
                {
                    MessageBox.Show("Orden no existe");
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT * FROM empresasipl WHERE id =" + empresa + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT id,cp,nif FROM clientesipl WHERE id =" + idcli + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        cpcli = "" + row["cp"];
                        nifcli = "" + row["nif"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                if (currency == "USD") { csymbol = "\"US$\"#,##0.00;-\"US$\"#,##0.00"; }
                if (currency == "EUR") { csymbol = "\"\"#,##0.00€;-\"\"#,##0.00€"; }
                if (currency == "") { csymbol = "\"\"#,##0.00;-\"\"#,##0.00"; }

                proforma.Cells["A1:I1"].Merge = true;
                proforma.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A1"].Style.Font.Size = 16;
                proforma.Cells["A1"].Style.Font.Name = "Calibri";
                proforma.Cells["A1"].Style.Font.Bold = true;
                proforma.Cells["A1"].Value = "" + nome;


                proforma.Cells["A2:I2"].Merge = true;
                proforma.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["A2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A2"].Style.Font.Size = 9;
                proforma.Cells["A2"].Style.Font.Name = "Calibri";
                proforma.Cells["A2"].Style.WrapText = true;
                proforma.Row(2).Height = 32;
                proforma.Cells["A2"].Value = "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe + "\r\n" + "CIF: " + nife;
                //proforma.Cells["A2"].Value = "CIF: " + nife; ;

                proforma.Cells["A3:I3"].Merge = true;
                proforma.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A3"].Style.Font.Size = 11;
                proforma.Cells["A3"].Style.Font.Name = "Calibri";
                proforma.Cells["A3"].Style.WrapText = true;
                proforma.Row(3).Height = 10;
                proforma.Cells["A3"].Value = "";



                proforma.Cells["A4:I4"].Merge = true;
                proforma.Cells["A4"].Style.Font.Bold = true;
                proforma.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["A4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A4"].Style.Font.Size = 16;
                proforma.Cells["A4"].Style.Font.Name = "Calibri";
                proforma.Cells["A4"].Style.WrapText = true;
                proforma.Row(4).Height = 20;
                proforma.Cells["A4"].Value = "FACTURA";

                /*  
                  proforma.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                  proforma.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                  proforma.Cells["A7"].Style.Font.Size = 10;
                  proforma.Cells["A7"].Style.Font.Name = "Calibri";
                  proforma.Cells["A7"].Style.WrapText = true;
                  proforma.Row(7).Height = 20;
                  proforma.Cells["A7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                  proforma.Cells["A7"].Value = "BUYER:";

                  proforma.Cells["A8:A11"].Merge = true;
                  proforma.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                  proforma.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                  proforma.Cells["A8"].Style.Font.Size = 10;
                  proforma.Cells["A8"].Style.Font.Name = "Calibri";
                  proforma.Cells["A8"].Style.WrapText = true;
                  proforma.Cells["A8:A11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                  proforma.Cells["A8"].Value = "ADDRESS:";
                  */

                proforma.Cells["A6:E11"].Style.Border.BorderAround(ExcelBorderStyle.Thick);

                proforma.Cells["A6:E6"].Merge = true;
                proforma.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A6"].Style.Font.Size = 10;
                proforma.Cells["A6"].Style.Font.Name = "Calibri";
                proforma.Cells["A6"].Style.WrapText = true;
                proforma.Cells["A6"].Value = "" + nomcli;

                proforma.Cells["A7:E7"].Merge = true;
                proforma.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A7"].Style.Font.Size = 10;
                proforma.Cells["A7"].Style.Font.Name = "Calibri";
                proforma.Cells["A7"].Style.WrapText = true;
                proforma.Cells["A7"].Value = "" + nifcli;

                proforma.Cells["A8:E8"].Merge = true;
                proforma.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A8"].Style.Font.Size = 10;
                proforma.Cells["A8"].Style.Font.Name = "Calibri";
                proforma.Cells["A8"].Style.WrapText = true;
                proforma.Cells["A8"].Value = "" + callecli + " " + numcli + " " + numclii;

                proforma.Cells["A9:E9"].Merge = true;
                proforma.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A9"].Style.Font.Size = 10;
                proforma.Cells["A9"].Style.Font.Name = "Calibri";
                proforma.Cells["A9"].Style.WrapText = true;
                proforma.Cells["A9"].Value = "" + colcli + "";

                proforma.Cells["A10:E10"].Merge = true;
                proforma.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A10"].Style.Font.Size = 10;
                proforma.Cells["A10"].Style.Font.Name = "Calibri";
                proforma.Cells["A10"].Style.WrapText = true;
                proforma.Cells["A10"].Value = "" + muncli + ", " + edocli;

                proforma.Cells["A11:E11"].Merge = true;
                proforma.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A11"].Style.Font.Size = 10;
                proforma.Cells["A11"].Style.Font.Name = "Calibri";
                proforma.Cells["A11"].Style.WrapText = true;
                proforma.Cells["A11"].Value = "" + paiscli + ", CP: " + cpcli;

                var usCulture = new System.Globalization.CultureInfo("en-US");
                //DateTime result = DateTime.ParseExact("" + fecha.Replace("/","-"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                string iDate = "" + fecha;
                DateTime oDate = Convert.ToDateTime(iDate);

                proforma.Cells["H6:I8"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                proforma.Cells["H6"].Style.Font.Size = 10;
                proforma.Cells["H6"].Style.Font.Name = "Calibri";
                proforma.Cells["H6"].Value = "FECHA: ";

                proforma.Cells["I6"].Style.Font.Size = 10;
                proforma.Cells["I6"].Style.Font.Name = "Calibri";
                proforma.Cells["I6"].Value = "" + oDate.ToString("dd-MM-yyyy");


                proforma.Cells["H7"].Style.Font.Size = 10;
                proforma.Cells["H7"].Style.Font.Name = "Calibri";
                proforma.Cells["H7"].Value = "Nº:";

                proforma.Cells["I7"].Style.Font.Size = 10;
                proforma.Cells["I7"].Style.Font.Name = "Calibri";
                proforma.Cells["I7"].Value = "" + folio;

                proforma.Cells["H8"].Style.Font.Size = 10;
                proforma.Cells["H8"].Style.Font.Name = "Calibri";
                proforma.Cells["H8"].Value = "FACTURA Nº:";

                proforma.Cells["I8"].Style.Font.Size = 10;
                proforma.Cells["I8"].Style.Font.Name = "Calibri";
                proforma.Cells["I8"].Value = "" + numee;



                proforma.Cells["G10:I10"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                proforma.Cells["G10:H10"].Merge = true;
                proforma.Cells["G10"].Style.Font.Size = 10;
                proforma.Cells["G10"].Style.Font.Name = "Calibri";
                proforma.Cells["G10"].Value = "PARTIDA ESTADISTICA:";

                proforma.Cells["I10"].Style.Font.Size = 10;
                proforma.Cells["I10"].Style.Font.Name = "Calibri";
                proforma.Cells["I10"].Value = "" + albaran;


                /*
                string strDate = "" + fecha;
                DateTimeFormatInfo dtinfo = new DateTimeFormatInfo();
                dtinfo.ShortDatePattern = "dd,MMM,yyyy";
                DateTime resultDate = Convert.ToDateTime(strDate, dtinfo);
                */

                //proforma.Cells["A13"].Style.Font.Size = 10;
                //proforma.Cells["A13"].Style.Font.Name = "Calibri";
                //proforma.Cells["A13"].Value = "Per S.S. " + origdest;

                //primeros 4 encabezados
                proforma.Cells["A13:I13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                proforma.Cells["A13:I13"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(170, 170, 170));


                proforma.Cells["A13:C13"].Merge = true;
                proforma.Cells["A13:C13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["A13"].Style.Font.Bold = true;
                proforma.Cells["A13"].Style.Font.Size = 11;
                proforma.Cells["A13"].Style.Font.Name = "Calibri";
                proforma.Cells["A13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["A13"].Value = "DESCRIPCIÓN DE LA MERCANCÍA";


                proforma.Cells["D13:G13"].Merge = true;
                proforma.Cells["D13:G13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["D13"].Style.Font.Bold = true;
                proforma.Cells["D13"].Style.Font.Size = 11;
                proforma.Cells["D13"].Style.Font.Name = "Calibri";
                proforma.Cells["D13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["D13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["D13"].Value = "UNIDADES";



                proforma.Cells["H13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["H13"].Style.Font.Size = 11;
                proforma.Cells["H13"].Style.Font.Bold = true;
                proforma.Cells["H13"].Style.Font.Name = "Calibri";
                proforma.Cells["H13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["H13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["H13"].Value = "PRECIO NETO";


                proforma.Cells["I13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["I13"].Style.Font.Size = 11;
                proforma.Cells["I13"].Style.Font.Bold = true;
                proforma.Cells["I13"].Style.Font.Name = "Calibri";
                proforma.Cells["I13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["I13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["I13"].Value = "IMPORTE TOTAL";
                // fin de 4 headers


                proforma.Cells["A14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["A14"].Style.Font.Size = 10;
                proforma.Cells["A14"].Style.Font.Name = "Calibri";
                proforma.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A14"].Style.WrapText = true;
                proforma.Cells["A14"].Value = "FORMATO";


                proforma.Cells["B14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["B14"].Style.Font.Size = 10;
                proforma.Cells["B14"].Style.Font.Name = "Calibri";
                proforma.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["B14"].Style.WrapText = true;
                proforma.Cells["B14"].Value = "MODELO";
                /*
                proforma.Cells["B14:B15"].Merge = true;
                proforma.Cells["B14:B15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["B14"].Style.Font.Size = 10;
                proforma.Cells["B14"].Style.Font.Name = "Calibri";
                proforma.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["B14"].Style.WrapText = true;
                proforma.Cells["B14"].Value = "MODEL";
                */

                proforma.Cells["C14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["C14"].Style.Font.Size = 10;
                proforma.Cells["C14"].Style.Font.Name = "Calibri";
                proforma.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["C14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["C14"].Style.WrapText = true;
                proforma.Cells["C14"].Value = "CLASE";


                proforma.Cells["D14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["D14"].Style.Font.Size = 10;
                proforma.Cells["D14"].Style.Font.Name = "Calibri";
                proforma.Cells["D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["D14"].Style.WrapText = true;
                proforma.Cells["D14"].Value = "PALLETS";


                proforma.Cells["E14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["E14"].Style.Font.Size = 10;
                proforma.Cells["E14"].Style.Font.Name = "Calibri";
                proforma.Cells["E14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["E14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["E14"].Style.WrapText = true;
                proforma.Cells["E14"].Value = "CAJAS";

                proforma.Cells["F14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["F14"].Style.Font.Size = 10;
                proforma.Cells["F14"].Style.Font.Name = "Calibri";
                proforma.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["F14"].Style.WrapText = true;
                proforma.Cells["F14"].Value = "M²";


                proforma.Cells["G14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["G14"].Style.Font.Size = 10;
                proforma.Cells["G14"].Style.Font.Name = "Calibri";
                proforma.Cells["G14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["G14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["G14"].Style.WrapText = true;
                proforma.Cells["G14"].Value = "PIEZAS";

                proforma.Cells["H14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["H14"].Style.Font.Size = 10;
                proforma.Cells["H14"].Style.Font.Name = "Calibri";
                proforma.Cells["H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["H14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["H14"].Style.WrapText = true;
                proforma.Cells["H14"].Value = "" + currency;

                proforma.Cells["I14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["I14"].Style.Font.Size = 10;
                proforma.Cells["I14"].Style.Font.Name = "Calibri";
                proforma.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["I14"].Style.WrapText = true;
                proforma.Cells["I14"].Value = "" + currency;
                /*
                proforma.Cells["I14:I15"].Merge = true;
                proforma.Cells["I14:I15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["I14"].Style.Font.Size = 10;
                proforma.Cells["I14"].Style.Font.Name = "Calibri";
                proforma.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["I14"].Style.WrapText = true;
                proforma.Cells["I14"].Value = "No.of CTNRS";
                */


                #endregion headerproform

                #region builerproform
                int rownu = 15, contu = 1, rowni = 14;
                string caja = "", mtscaja = "", kgspiece = "", kgscaja = "", sku = "";
                kgscontainer = 0; kgsacums = 0;


                double ctns = 0, sqm = 0, tctns = 0, tsqm = 0, cajad = 0, mtsd = 0, pallets = 0, pieces = 0;
                string size = "", sizel = "", pallet = "", units = "", cants = "";
                query = "SELECT * FROM rowsipl WHERE ord ='" + id + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {

                        string querya = "SELECT * FROM artsipl WHERE clave ='" + row["clave"] + "';";
                        SqlCommand cma = new SqlCommand(querya, con);
                        SqlDataAdapter daa = new SqlDataAdapter(cma);
                        DataTable dta = new DataTable();
                        daa.Fill(dta);
                        int cuentaa = dta.Rows.Count;
                        if (cuentaa > 0)
                        {
                            foreach (DataRow rowa in dta.Rows)
                            {
                                size = "" + rowa["size"]; sizel = "" + rowa["sizel"];
                                caja = "" + rowa["caja"]; mtscaja = "" + rowa["mtscaja"]; kgspiece = "" + rowa["kgspiece"];
                                kgscaja = "" + rowa["kgscaja"]; sku = "" + rowa["size"];
                                pallet = "" + rowa["pallet"];
                                units = "" + rowa["ume"];
                            }

                        }
                        cma.Dispose(); daa.Dispose(); dta.Dispose();



                        //ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                        //ord, cant, clave, ume, pu, importe, container
                        proforma.Cells["A" + rownu].Merge = true;
                        proforma.Cells["A" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.Font.Size = 9;
                        proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["A" + rownu].Style.WrapText = true;
                        proforma.Cells["A" + rownu].Value = "" + size + "X" + sizel;

                        proforma.Cells["B" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["B" + rownu].Style.Font.Size = 9;
                        proforma.Cells["B" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["B" + rownu].Style.WrapText = true;
                        proforma.Cells["B" + rownu].Value = "" + row["clave"];

                        proforma.Cells["C" + rownu].Merge = true;
                        proforma.Cells["C" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["C" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["C" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["C" + rownu].Style.Font.Size = 9;
                        proforma.Cells["C" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["C" + rownu].Style.WrapText = true;
                        proforma.Cells["C" + rownu].Value = 1;

                        /*
                        proforma.Cells["B" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["B" + rownu].Style.Font.Size = 9;
                        proforma.Cells["B" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["B" + rownu].Style.WrapText = true;
                        proforma.Cells["B" + rownu].Value ="" + row["clave"];
                        */


                        //proforma.Cells["C" + rownu].Value = "" + units; (M2 OR PIECE)


                        sqm = 0;
                        tctns = 0;
                        tsqm = 0;
                        cajad = 0;
                        mtsd = 0;
                        pallets = 0;
                        pieces = 0;

                        try { ctns = double.Parse("" + caja); }
                        catch { ctns = 0; }

                        try { sqm = double.Parse("" + mtscaja); }
                        catch { sqm = 0; }

                        try { pallets = double.Parse("" + pallet); }
                        catch { pallets = 0; }


                        try {
                            pieces = (ctns * double.Parse("" + row["cant"])) / sqm;
                        }
                        catch {
                            pieces = 0;
                        }


                        try
                        {
                            tctns = double.Parse("" + row["cant"]) / ctns;
                        }
                        catch { tctns = 0; }

                        try
                        {
                            tsqm = sqm * tctns;
                        }
                        catch { tsqm = 0; }
                        if (row["pallets"].ToString() != "")
                        {
                            pallets = double.Parse("" + row["pallets"]);
                        }
                        else
                        {
                            try
                            {
                                pallets = double.Parse("" + row["cant"]) / pallets;
                            }
                            catch { pallets = 1; }
                            //if (pallets < 1) { pallets = 1; }
                            if (pallets.ToString() == "∞") { pallets = 0; }
                        }
                        try {
                            cajad = double.Parse("" + row["cant"]) / sqm;
                        }
                        catch { cajad = 0; }



                        proforma.Cells["D" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["D" + rownu].Style.Font.Size = 9;
                        proforma.Cells["D" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["D" + rownu].Style.WrapText = true;
                        proforma.Cells["D" + rownu].Style.Numberformat.Format = "#,##0.00";
                        proforma.Cells["D" + rownu].Value = pallets;
                        //proforma.Cells["D" + rownu].Value = double.Parse("" + row["cant"]);

                        proforma.Cells["E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["E" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["E" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["E" + rownu].Style.Font.Size = 9;
                        proforma.Cells["E" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["E" + rownu].Style.WrapText = true;
                        proforma.Cells["E" + rownu].Style.Numberformat.Format = "#,##0";
                        proforma.Cells["E" + rownu].Value = double.Parse("" + cajad);

                        proforma.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["F" + rownu].Style.Font.Size = 9;
                        proforma.Cells["F" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["F" + rownu].Style.WrapText = true;
                        proforma.Cells["F" + rownu].Style.Numberformat.Format = "#,##0.00";
                        proforma.Cells["F" + rownu].Value = double.Parse("" + row["cant"]);

                        proforma.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["G" + rownu].Style.Font.Size = 9;
                        proforma.Cells["G" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["G" + rownu].Style.WrapText = true;
                        proforma.Cells["G" + rownu].Style.Numberformat.Format = "#,##0";
                        proforma.Cells["G" + rownu].Value = pieces;
                        //proforma.Cells["G" + rownu].Value = double.Parse("" + row["pu"]);


                        proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["H" + rownu].Style.Font.Size = 9;
                        proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["H" + rownu].Style.WrapText = true;
                        proforma.Cells["H" + rownu].Style.Numberformat.Format = "" + csymbol;
                        proforma.Cells["H" + rownu].Value = double.Parse("" + row["pu"]);


                        proforma.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["I" + rownu].Style.Font.Size = 9;
                        proforma.Cells["I" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["I" + rownu].Style.WrapText = true;
                        proforma.Cells["I" + rownu].Style.Numberformat.Format = "" + csymbol;
                        proforma.Cells["I" + rownu].Value = double.Parse("" + row["importe"]);

                        rownu = rownu + 1; contu = contu + 1;

                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                //VOLCAR SERVICIOS EN LA FACTURA

                query = "SELECT * FROM rowsservpl WHERE ord ='" + id + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                int cuentasrv = dt.Rows.Count;
                if (cuentasrv > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {

                        sqm = 0;
                        tctns = 0;
                        tsqm = 0;
                        cajad = 0;
                        mtsd = 0;
                        pallets = 0;
                        ctns = 0;


                        try { ctns = double.Parse("" + row["cant"]); }
                        catch { ctns = 0; }

                        try { sqm = double.Parse("" + row["cu"]); }
                        catch { sqm = 0; }

                        try { tsqm = double.Parse("" + row["total"]); }
                        catch { tsqm = 0; }


                        /*
                        proforma.Cells["A" + rownu].Merge = true;
                        proforma.Cells["A" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.Font.Size = 9;
                        proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["A" + rownu].Style.WrapText = true;
                        proforma.Cells["A" + rownu].Value = "" + row["clave"];
                        */
                        proforma.Cells["A" + rownu + ":F" + rownu].Merge = true;
                        proforma.Cells["A" + rownu + ":F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.Font.Size = 9;
                        proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["A" + rownu].Style.WrapText = true;
                        proforma.Cells["A" + rownu].Value = "" + row["descrip"];

                        proforma.Cells["G" + rownu].Merge = true;
                        proforma.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["G" + rownu].Style.Font.Size = 9;
                        proforma.Cells["G" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["G" + rownu].Style.WrapText = true;
                        proforma.Cells["G" + rownu].Value = ctns;




                        proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["H" + rownu].Style.Font.Size = 9;
                        proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["H" + rownu].Style.WrapText = true;
                        proforma.Cells["H" + rownu].Style.Numberformat.Format = "" + csymbol;
                        //proforma.Cells["H" + rownu].Style.Numberformat.Format = "\"US$\"#,##0.00;-\"US$\"#,##0.00";
                        proforma.Cells["H" + rownu].Value = sqm;


                        proforma.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["I" + rownu].Style.Font.Size = 9;
                        proforma.Cells["I" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["I" + rownu].Style.WrapText = true;
                        proforma.Cells["I" + rownu].Style.Numberformat.Format = "" + csymbol;
                        proforma.Cells["I" + rownu].Value = tsqm;

                        rownu = rownu + 1; contu = contu + 1;

                    }
                }
                else
                {

                    sqm = 0;
                    tctns = 0;
                    tsqm = 0;
                    cajad = 0;
                    mtsd = 0;
                    pallets = 0;
                    ctns = 0;

                    /*
                    proforma.Cells["A" + rownu].Merge = true;
                    proforma.Cells["A" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["A" + rownu].Style.Font.Size = 9;
                    proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                    proforma.Cells["A" + rownu].Style.WrapText = true;
                    proforma.Cells["A" + rownu].Value = "" + row["clave"];
                    */
                    proforma.Cells["A" + rownu + ":F" + rownu].Merge = true;
                    //proforma.Cells["A" + rownu + ":F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["A" + rownu].Style.Font.Size = 9;
                    proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                    proforma.Cells["A" + rownu].Style.WrapText = true;
                    proforma.Cells["A" + rownu].Value = "";

                    proforma.Cells["G" + rownu].Merge = true;
                    //proforma.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    proforma.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["G" + rownu].Style.Font.Size = 9;
                    proforma.Cells["G" + rownu].Style.Font.Name = "Calibri";
                    proforma.Cells["G" + rownu].Style.WrapText = true;
                    //proforma.Cells["G" + rownu].Value = ctns;




                    //proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["H" + rownu].Style.Font.Size = 9;
                    proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                    proforma.Cells["H" + rownu].Style.WrapText = true;
                    proforma.Cells["H" + rownu].Style.Numberformat.Format = "" + csymbol;
                    //proforma.Cells["H" + rownu].Value = sqm;


                    //proforma.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    proforma.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["I" + rownu].Style.Font.Size = 9;
                    proforma.Cells["I" + rownu].Style.Font.Name = "Calibri";
                    proforma.Cells["I" + rownu].Style.WrapText = true;
                    proforma.Cells["I" + rownu].Style.Numberformat.Format = "" + csymbol;
                    //proforma.Cells["I" + rownu].Value = tsqm;
                    cuentasrv = 1;
                    rownu = rownu + 1; contu = contu + 1;

                }
                cm.Dispose(); da.Dispose(); dt.Dispose();




                #endregion builerproform

                #region footerproform


                int rowf = rownu;
                if (rowf > rowni)
                {
                    //rowf = rowf - 1;
                }
                rownu = rownu + 1;

                proforma.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["G" + rownu].Style.Font.Size = 9;
                proforma.Cells["G" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["G" + rownu].Style.WrapText = true;
                proforma.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["H" + rownu].Style.Font.Size = 9;
                proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["H" + rownu].Style.WrapText = true;
                proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["H" + rownu].Style.Numberformat.Format = "" + csymbol;
                proforma.Cells["H" + rownu].Formula = "=SUM(H" + rowni + ":H" + rowf + ")";





                proforma.Cells["D" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["D" + rownu].Style.Font.Size = 9;
                proforma.Cells["D" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["D" + rownu].Style.WrapText = true;
                proforma.Cells["D" + rownu].Value = "PALLETS";


                proforma.Cells["E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["E" + rownu].Style.Font.Size = 9;
                proforma.Cells["E" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["E" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["E" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["E" + rownu].Style.WrapText = true;
                proforma.Cells["E" + rownu].Value = "CAJAS";

                proforma.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["F" + rownu].Style.Font.Size = 9;
                proforma.Cells["F" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["F" + rownu].Style.WrapText = true;
                proforma.Cells["F" + rownu].Value = "M²";


                proforma.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["G" + rownu].Style.Font.Size = 9;
                proforma.Cells["G" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["G" + rownu].Style.WrapText = true;
                proforma.Cells["G" + rownu].Value = "PIEZAS";

                proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["H" + rownu].Style.Font.Size = 9;
                proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["H" + rownu].Style.WrapText = true;
                proforma.Cells["H" + rownu].Value = "PESO NETO";

                proforma.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["I" + rownu].Style.Font.Size = 9;
                proforma.Cells["I" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["I" + rownu].Style.WrapText = true;
                proforma.Cells["I" + rownu].Value = "PESO BRUTO";


                rownu = rownu + 1;



                proforma.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["D" + rownu].Style.Font.Size = 9;
                proforma.Cells["D" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["D" + rownu].Style.WrapText = true;
                proforma.Cells["D" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["D" + rownu].Style.Numberformat.Format = "##0.00";
                proforma.Cells["D" + rownu].Formula = "=SUM(D" + rowni + ":D" + ((rowf - cuentasrv) - 1) + ")";


                proforma.Cells["E" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["E" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["E" + rownu].Style.Font.Size = 9;
                proforma.Cells["E" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["E" + rownu].Style.WrapText = true;
                proforma.Cells["E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["E" + rownu].Style.Numberformat.Format = "##0";
                proforma.Cells["E" + rownu].Formula = "=SUM(E" + rowni + ":E" + ((rowf - cuentasrv) - 1) + ")";

                proforma.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["F" + rownu].Style.Font.Size = 9;
                proforma.Cells["F" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["F" + rownu].Style.WrapText = true;
                proforma.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["F" + rownu].Style.Numberformat.Format = "##0.00";
                proforma.Cells["F" + rownu].Formula = "=SUM(F" + rowni + ":F" + ((rowf - cuentasrv) - 1) + ")";

                proforma.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["G" + rownu].Style.Font.Size = 9;
                proforma.Cells["G" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["G" + rownu].Style.WrapText = true;
                proforma.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["G" + rownu].Style.Numberformat.Format = "##0";
                proforma.Cells["G" + rownu].Formula = "=SUM(G" + rowni + ":G" + ((rowf - cuentasrv)-1) + ")";



                proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["H" + rownu].Style.Font.Size = 9;
                proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["H" + rownu].Style.WrapText = true;
                proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["H" + rownu].Style.Numberformat.Format = "##0.00";
                proforma.Cells["H" + rownu].Value = double.Parse("" + netofact.Replace(",",""));

                cellrow1 = "H" + rownu;

                proforma.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["I" + rownu].Style.Font.Size = 9;
                proforma.Cells["I" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["I" + rownu].Style.WrapText = true;
                proforma.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["I" + rownu].Style.Numberformat.Format = "##0.00";
                proforma.Cells["I" + rownu].Value = double.Parse("" + brutosfact);

                cellrow2 = "I" + rownu;

                rownu = rownu + 2;
                /*
                proforma.Cells["B" + rownu + ":G" + rownu].Merge = true;
                proforma.Cells["B" + rownu + ":G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["B" + rownu].Style.Font.Size = 9;
                proforma.Cells["B" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                proforma.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["B" + rownu].Value = "KGS: " + albaran ;
                */


                //CICLO PARA SACAR LOS NUMROS DE CONTENEDOR
                rownu = rownu + 2;

                proforma.Cells["B" + rownu + ":C" + rownu].Merge = true;
                proforma.Cells["B" + rownu + ":C" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                proforma.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["B" + rownu].Style.Font.Size = 9;
                proforma.Cells["B" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["B" + rownu].Style.WrapText = true;
                proforma.Cells["B" + rownu].Value = "CONTENEDOR(ES)";

                proforma.Cells["D" + rownu + ":E" + rownu].Merge = true;
                proforma.Cells["D" + rownu + ":E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                proforma.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["D" + rownu].Style.Font.Size = 9;
                proforma.Cells["D" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["D" + rownu].Style.WrapText = true;
                proforma.Cells["D" + rownu].Value = "PRERECINTO";

                proforma.Cells["F" + rownu + ":G" + rownu].Merge = true;
                proforma.Cells["F" + rownu + ":G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                proforma.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["F" + rownu].Style.Font.Size = 9;
                proforma.Cells["F" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["F" + rownu].Style.WrapText = true;
                proforma.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["F" + rownu].Value = "PESO NETO EN KGS";


                proforma.Cells["H" + rownu + ":I" + rownu].Merge = true;
                proforma.Cells["H" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["H" + rownu].Style.Font.Size = 9;
                proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["H" + rownu].Style.WrapText = true;
                proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["H" + rownu].Value = "PESO BRUTO EN KGS";



                rownu = rownu + 1;

                //GENERAR CALCULO DE TOTALES EN PESO POR CONTENDOR

                //CALCULAR TOTAL DE CONTENEDOR EN KILOS EN BASE A LAS PIEZAS O METROS
                //ord, cant, clave, ume, pu, importe, container
                string kgsacumulados = "";
                mtspartidas = 0; mtsacum = 0; kgs1caja = 0; mts1caja = 0; tonelaje = 0;
                tarimaskgs = 0; kgstar = 0; kilosXpallets = 0;
                kgstar = double.Parse("" + palletkgs.Text.ToString());
                string queryc = "SELECT * FROM containersipl WHERE ord ='" + id + "' ORDER BY id ASC;";
                SqlCommand cmc = new SqlCommand(queryc, con);
                SqlDataAdapter dac = new SqlDataAdapter(cmc);
                DataTable dtc = new DataTable();
                dac.Fill(dtc);
                int cuentac = dtc.Rows.Count;
                if (cuentac > 0)
                {
                    foreach (DataRow rowc in dtc.Rows)
                    {


                        /*
                        string queryckg = @"SELECT r.ord,r.cant,r.clave,r.pallets, " +
                        "ISNULL((SELECT  caja FROM artsipl WHERE clave = r.clave),0) as caja, " +
                        "ISNULL((SELECT  mtscaja FROM artsipl WHERE clave = r.clave),0) as mtscaja, " +
                        "ISNULL((SELECT  kgscaja FROM artsipl WHERE clave = r.clave),0) as kgscaja, " +
                        "ISNULL((SELECT  pallet FROM artsipl WHERE clave = r.clave),0) as pallet "+
                        "FROM rowsipl  AS r " +
                        "WHERE ord ='" + id + "' AND container='" + rowc["container"] + "' ORDER BY id ASC;";

                        SqlCommand cmckg = new SqlCommand(queryckg, con);
                        SqlDataAdapter dackg = new SqlDataAdapter(cmckg);
                        DataTable dtckg = new DataTable();
                        dackg.Fill(dtckg);
                        int cuentackg = dtckg.Rows.Count;
                        if (cuentackg > 0)
                        {
                            foreach (DataRow rowckg in dtckg.Rows)
                            {
                                kgscaja = ""; mtscaja = "";
                                 try
                                { mtspartidas = double.Parse("" + rowckg["cant"]); }
                                catch
                                { mtspartidas = 0; }

                                mtscaja = "" + rowckg["mtscaja"];
                                kgscaja = "" + rowckg["kgscaja"];

                                try { mts1caja = double.Parse("" + mtscaja); }
                                catch { mts1caja = 0; }

                                try { kgs1caja = double.Parse("" + kgscaja); }
                                catch { kgs1caja = 0; }
                                
                                //convertir los metros a Kilos
                                tonelaje = 0;
                                tonelaje = (1 * mtspartidas) / mts1caja;
                                tonelaje = kgs1caja * tonelaje;
                                tonelajecontainer = tonelajecontainer + tonelaje;
                                kilosXpallets = kgstar * double.Parse("" + rowckg["pallets"]);
                                
                                //tarimaskgs = tarimaskgs + kilosXpallets;                                
                                tarimaskgs = tarimaskgs + kilosXpallets;
                                //MessageBox.Show(" CLAVE: " + rowckg["clave"] +  " TOTAL CONTAINER: " + tonelajecontainer +  " KILOS PALLET:" + kilosXpallets + " ACUM TARIMAS: " + tarimaskgs);
                            }
                        } else  {
                            tonelaje = 0;
                            //tarimaskgs = 0;
                            kilosXpallets = 0;
                        }
                        cmckg.Dispose();dackg.Dispose();dtckg.Dispose();


                      /* 
                       string queryckg2 = @"//SELECT sum(CONVERT(decimal, r.pallets)) as Cuenta FROM rowsipl  AS r WHERE r.ord = '14' ;";
                        SqlCommand cmckg2 = new SqlCommand(queryckg2, con);
                        SqlDataAdapter dackg2 = new SqlDataAdapter(cmckg);
                        DataTable dtckg2 = new DataTable();
                        dackg2.Fill(dtckg2);
                        int cuentackg2 = dtckg2.Rows.Count;
                        if (cuentackg2 > 0)
                        {
                            foreach (DataRow rowckg2 in dtckg2.Rows)
                            {

                            }

                        }
                        */
                        proforma.Cells["B" + rownu + ":C" + rownu].Merge = true;
                        proforma.Cells["B" + rownu + ":C" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        proforma.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["B" + rownu].Style.Font.Size = 9;
                        proforma.Cells["B" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["B" + rownu].Style.WrapText = true;
                        proforma.Cells["B" + rownu].Value = "" + rowc["container"];

                        proforma.Cells["D" + rownu + ":E" + rownu].Merge = true;
                        proforma.Cells["D" + rownu + ":E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        proforma.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["D" + rownu].Style.Font.Size = 9;
                        proforma.Cells["D" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["D" + rownu].Style.WrapText = true;
                        proforma.Cells["D" + rownu].Value = "" + rowc["precinto"];

                        proforma.Cells["F" + rownu + ":G" + rownu].Merge = true;
                        proforma.Cells["F" + rownu + ":G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        proforma.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["F" + rownu].Style.Font.Size = 9;
                        proforma.Cells["F" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["F" + rownu].Style.WrapText = true;
                        proforma.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["F" + rownu].Style.Numberformat.Format = "#,##0.00";
                        proforma.Cells["F" + rownu].Value = double.Parse("" + rowc["pesoneto"]);


                        proforma.Cells["H" + rownu + ":I" + rownu].Merge = true;
                        proforma.Cells["H" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        proforma.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["H" + rownu].Style.Font.Size = 9;
                        proforma.Cells["H" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["H" + rownu].Style.WrapText = true;
                        proforma.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["H" + rownu].Style.Numberformat.Format = "#,##0.00";
                        proforma.Cells["H" + rownu].Value = double.Parse("" + rowc["pesobruto"]);
                        rownu = rownu + 1;
                        /*
                        //acumular PESOS
                        netofull = netofull + tonelajecontainer;
                        brutofull = brutofull + (tonelajecontainer + tarimaskgs);
                        //fin de acumular el pesado
                        tarimaskgs = 0;
                        tonelajecontainer = 0;

                        */

                    }
                }
                cmc.Dispose();dac.Dispose(); dtc.Dispose();
                //FIN DE CICLO PARA SACAR LOS NUMEROS DE CONTENEDOR

                // FIN DE CALCULO DE PESOS CONTENEDOR


                //SUMAS TOTALES DE BRUTO Y NETO
                //proforma.Cells["" + cellrow1].Style.Numberformat.Format = "#,##0.00";
                //proforma.Cells["" + cellrow1].Value = double.Parse("" + netofact);
                //proforma.Cells["" + cellrow2].Style.Numberformat.Format = "#,##0.00";
                //proforma.Cells["" + cellrow2].Value = double.Parse("" + brutosfact);

                //FIN DE SUMATORIAS DE BRUTO Y NETO
                proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["A" + rownu].Style.Font.Size = 9;
                proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["A" + rownu].Style.WrapText = true;
                proforma.Cells["A" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["A" + rownu].Value = "DESTINO:";

                proforma.Cells["B" + rownu + ":I" + rownu].Merge = true;
                proforma.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                proforma.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["B" + rownu].Style.Font.Size = 9;
                proforma.Cells["B" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["B" + rownu].Style.WrapText = true;
                proforma.Cells["B" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["B" + rownu].Value = "" + origdest;





                rownu = rownu + 1;


                proforma.Cells["G" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["G" + rownu + ":H" + rownu].Merge = true;
                proforma.Cells["G" + rownu + ":H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["G" + rownu + ":I" + rownu].Style.Font.Bold = true;
                proforma.Cells["G" + rownu + ":I" + rownu].Style.Font.Size = 9;
                proforma.Cells["G" + rownu + ":I" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["G" + rownu + ":I" + rownu].Style.Fill.PatternType = ExcelFillStyle.Solid;
                proforma.Cells["G" + rownu + ":I" + rownu].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(170, 170, 170));

                proforma.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                proforma.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["G" + rownu].Style.Font.Size = 9;
                proforma.Cells["G" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["G" + rownu].Style.WrapText = true;
                proforma.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["G" + rownu].Value = "TOTAL FACTURA";


                proforma.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                proforma.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                proforma.Cells["I" + rownu].Style.Font.Size = 9;
                proforma.Cells["I" + rownu].Style.Font.Name = "Calibri";
                proforma.Cells["I" + rownu].Style.WrapText = true;
                proforma.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                proforma.Cells["I" + rownu].Style.Numberformat.Format = "" + csymbol;
                proforma.Cells["I" + rownu].Formula = "=SUM(I" + rowni + ":I" + rowf + ")";

                rownu = rownu + 1;

                if (obs1 != "")
                {
                    int rowes = 32;
                    rownu = rownu + 2;
                    proforma.Cells["A" + rownu + ":I" + rownu].Merge = true;
                    proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["A" + rownu].Style.Font.Size = 9;
                    proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                    proforma.Cells["A" + rownu].Style.Font.Bold = true;
                    proforma.Cells["A" + rownu].Style.WrapText = true;
                    proforma.Row(rownu).Height = 15;
                    proforma.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    proforma.Cells["A" + rownu].Value = "OBSERVACIONES:";

                    rownu = rownu + 1;
                    proforma.Cells["A" + rownu + ":I" + rownu].Merge = true;
                    proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["A" + rownu].Style.Font.Size = 9;
                    proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                    proforma.Cells["A" + rownu].Style.WrapText = true;
                    proforma.Row(rownu).Height = rowes;
                    proforma.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    proforma.Cells["A" + rownu].Value = "" + obs1;

                    if (obs2 != "")
                    {
                        rownu = rownu + 1;
                        proforma.Cells["A" + rownu + ":I" + rownu].Merge = true;
                        proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.Font.Size = 9;
                        proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["A" + rownu].Style.WrapText = true;
                        proforma.Row(rownu).Height = rowes;
                        proforma.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["A" + rownu].Value = "" + obs2;
                    }

                    if (obs3 != "")
                    {
                        rownu = rownu + 1;
                        proforma.Cells["A" + rownu + ":I" + rownu].Merge = true;
                        proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.Font.Size = 9;
                        proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["A" + rownu].Style.WrapText = true;
                        proforma.Row(rownu).Height = rowes;
                        proforma.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["A" + rownu].Value = "" + obs3;
                    }

                    if (obs4 != "")
                    {
                        rownu = rownu + 1;
                        proforma.Cells["A" + rownu + ":I" + rownu].Merge = true;
                        proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.Font.Size = 9;
                        proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["A" + rownu].Style.WrapText = true;
                        proforma.Row(rownu).Height = rowes;
                        proforma.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["A" + rownu].Value = "" + obs4;
                    }
                    if (obs5 != "")
                    {
                        rownu = rownu + 1;
                        proforma.Cells["A" + rownu + ":I" + rownu].Merge = true;
                        proforma.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        proforma.Cells["A" + rownu].Style.Font.Size = 9;
                        proforma.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        proforma.Cells["A" + rownu].Style.WrapText = true;
                        proforma.Row(rownu).Height = rowes;
                        proforma.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        proforma.Cells["A" + rownu].Value = "" + obs5;
                    }
                }



                #endregion footerproform


                #region headerpackinglistv2
                /*
                concentrado.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                concentrado.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                */

                con = new SqlConnection("" + config.cade);
                con.Open();
                query = "SELECT * FROM invoicespl WHERE id =" + idinvo.Text + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {

                        id = "" + row["id"]; folio = "" + row["folio"]; empresa = "" + row["empresa"]; idcli = "" + row["idcli"];
                        nomcli = "" + row["nomcli"];
                        callecli = "" + row["callecli"];
                        numcli = "" + row["numcli"]; numclii = "" + row["numclii"]; colcli = "" + row["colcli"];
                        muncli = "" + row["muncli"];
                        edocli = "" + row["edocli"]; paiscli = "" + row["paiscli"];
                        fecha = "" + row["fecha"];
                        origdest = "" + row["origdest"]; tot = "" + row["tot"];
                        currency = "" + row["currency"];

                        obs1 = "" + row["obs1"];
                        obs2 = "" + row["obs2"];
                        obs3 = "" + row["obs3"];
                        obs4 = "" + row["obs4"];
                        obs5 = "" + row["obs5"];

                        albaran = "" + row["albaran"];
                        numee = "" + row["number"];
                        netofact = "" + row["pesoneto"];
                        brutosfact = "" + row["pesobruto"];
                    }
                }
                else
                {
                    MessageBox.Show("Orden no existe");
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT * FROM empresasipl WHERE id =" + empresa + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT id,cp,nif FROM clientesipl WHERE id =" + idcli + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        cpcli = "" + row["cp"];
                        nifcli = "" + row["nif"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                if (currency == "USD") { csymbol = "\"US$\"#,##0.00;-\"US$\"#,##0.00"; }
                if (currency == "EUR") { csymbol = "\"\"#,##0.00€;-\"\"#,##0.00€"; }


                plist2.Cells["A1:I1"].Merge = true;
                plist2.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A1"].Style.Font.Size = 16;
                plist2.Cells["A1"].Style.Font.Name = "Calibri";
                plist2.Cells["A1"].Style.Font.Bold = true;
                plist2.Cells["A1"].Value = "" + nome;


                plist2.Cells["A2:I2"].Merge = true;
                plist2.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["A2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A2"].Style.Font.Size = 9;
                plist2.Cells["A2"].Style.Font.Name = "Calibri";
                plist2.Cells["A2"].Style.WrapText = true;
                plist2.Row(2).Height = 32;
                plist2.Cells["A2"].Value = "" + callee + "  " + nume + " " + numie + " " + cole + " " + "" + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe + "\r\n" + "CIF: " + nife;
                //plist2.Cells["A2"].Value = "CIF: " + nife; ;

                plist2.Cells["A3:I3"].Merge = true;
                plist2.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A3"].Style.Font.Size = 11;
                plist2.Cells["A3"].Style.Font.Name = "Calibri";
                plist2.Cells["A3"].Style.WrapText = true;
                plist2.Row(3).Height = 10;
                plist2.Cells["A3"].Value = "";



                plist2.Cells["A4:I4"].Merge = true;
                plist2.Cells["A4"].Style.Font.Bold = true;
                plist2.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["A4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A4"].Style.Font.Size = 16;
                plist2.Cells["A4"].Style.Font.Name = "Calibri";
                plist2.Cells["A4"].Style.WrapText = true;
                plist2.Row(4).Height = 20;
                plist2.Cells["A4"].Value = "PACKING LIST";

                /*  
                  plist2.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                  plist2.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                  plist2.Cells["A7"].Style.Font.Size = 10;
                  plist2.Cells["A7"].Style.Font.Name = "Calibri";
                  plist2.Cells["A7"].Style.WrapText = true;
                  plist2.Row(7).Height = 20;
                  plist2.Cells["A7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                  plist2.Cells["A7"].Value = "BUYER:";

                  plist2.Cells["A8:A11"].Merge = true;
                  plist2.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                  plist2.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                  plist2.Cells["A8"].Style.Font.Size = 10;
                  plist2.Cells["A8"].Style.Font.Name = "Calibri";
                  plist2.Cells["A8"].Style.WrapText = true;
                  plist2.Cells["A8:A11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                  plist2.Cells["A8"].Value = "ADDRESS:";
                  */

                plist2.Cells["A6:E11"].Style.Border.BorderAround(ExcelBorderStyle.Thick);

                plist2.Cells["A6:E6"].Merge = true;
                plist2.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A6"].Style.Font.Size = 10;
                plist2.Cells["A6"].Style.Font.Name = "Calibri";
                plist2.Cells["A6"].Style.WrapText = true;
                plist2.Cells["A6"].Value = "" + nomcli;

                plist2.Cells["A7:E7"].Merge = true;
                plist2.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A7"].Style.Font.Size = 10;
                plist2.Cells["A7"].Style.Font.Name = "Calibri";
                plist2.Cells["A7"].Style.WrapText = true;
                plist2.Cells["A7"].Value = "" + nifcli;

                plist2.Cells["A8:E8"].Merge = true;
                plist2.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A8"].Style.Font.Size = 10;
                plist2.Cells["A8"].Style.Font.Name = "Calibri";
                plist2.Cells["A8"].Style.WrapText = true;
                plist2.Cells["A8"].Value = "" + callecli + " " + numcli + " " + numclii;

                plist2.Cells["A9:E9"].Merge = true;
                plist2.Cells["A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A9"].Style.Font.Size = 10;
                plist2.Cells["A9"].Style.Font.Name = "Calibri";
                plist2.Cells["A9"].Style.WrapText = true;
                plist2.Cells["A9"].Value = "" + colcli + "";

                plist2.Cells["A10:E10"].Merge = true;
                plist2.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A10"].Style.Font.Size = 10;
                plist2.Cells["A10"].Style.Font.Name = "Calibri";
                plist2.Cells["A10"].Style.WrapText = true;
                plist2.Cells["A10"].Value = "" + muncli + ", " + edocli;

                plist2.Cells["A11:E11"].Merge = true;
                plist2.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A11"].Style.Font.Size = 10;
                plist2.Cells["A11"].Style.Font.Name = "Calibri";
                plist2.Cells["A11"].Style.WrapText = true;
                plist2.Cells["A11"].Value = "" + paiscli + ", CP: " + cpcli;

               

                plist2.Cells["H6:I8"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                plist2.Cells["H6"].Style.Font.Size = 10;
                plist2.Cells["H6"].Style.Font.Name = "Calibri";
                plist2.Cells["H6"].Value = "FECHA: ";

                plist2.Cells["I6"].Style.Font.Size = 10;
                plist2.Cells["I6"].Style.Font.Name = "Calibri";
                plist2.Cells["I6"].Value = "" + oDate.ToString("dd-MM-yyyy");


                plist2.Cells["H7"].Style.Font.Size = 10;
                plist2.Cells["H7"].Style.Font.Name = "Calibri";
                plist2.Cells["H7"].Value = "Nº:";

                plist2.Cells["I7"].Style.Font.Size = 10;
                plist2.Cells["I7"].Style.Font.Name = "Calibri";
                plist2.Cells["I7"].Value = "" + folio;

                plist2.Cells["H8"].Style.Font.Size = 10;
                plist2.Cells["H8"].Style.Font.Name = "Calibri";
                plist2.Cells["H8"].Value = "PL Nº:";

                plist2.Cells["I8"].Style.Font.Size = 10;
                plist2.Cells["I8"].Style.Font.Name = "Calibri";
                plist2.Cells["I8"].Value = "" + numee;


                plist2.Cells["G10:I10"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                plist2.Cells["G10:H10"].Merge = true;
                plist2.Cells["G10"].Style.Font.Size = 10;
                plist2.Cells["G10"].Style.Font.Name = "Calibri";
                plist2.Cells["G10"].Value = "PARTIDA ESTADISTICA:";

                plist2.Cells["I10"].Style.Font.Size = 10;
                plist2.Cells["I10"].Style.Font.Name = "Calibri";
                plist2.Cells["I10"].Value = "" + albaran;


                /*
                string strDate = "" + fecha;
                DateTimeFormatInfo dtinfo = new DateTimeFormatInfo();
                dtinfo.ShortDatePattern = "dd,MMM,yyyy";
                DateTime resultDate = Convert.ToDateTime(strDate, dtinfo);
                */

                //plist2.Cells["A13"].Style.Font.Size = 10;
                //plist2.Cells["A13"].Style.Font.Name = "Calibri";
                //plist2.Cells["A13"].Value = "Per S.S. " + origdest;

                //primeros 4 encabezados
                plist2.Cells["A13:I13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                plist2.Cells["A13:I13"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(170, 170, 170));


                plist2.Cells["A13:C13"].Merge = true;
                plist2.Cells["A13:C13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["A13"].Style.Font.Bold = true;
                plist2.Cells["A13"].Style.Font.Size = 11;
                plist2.Cells["A13"].Style.Font.Name = "Calibri";
                plist2.Cells["A13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["A13"].Value = "DESCRIPCIÓN DE LA MERCANCÍA";


                plist2.Cells["D13:G13"].Merge = true;
                plist2.Cells["D13:G13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["D13"].Style.Font.Bold = true;
                plist2.Cells["D13"].Style.Font.Size = 11;
                plist2.Cells["D13"].Style.Font.Name = "Calibri";
                plist2.Cells["D13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["D13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["D13"].Value = "UNIDADES";



                plist2.Cells["H13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["H13"].Style.Font.Size = 11;
                plist2.Cells["H13"].Style.Font.Bold = true;
                plist2.Cells["H13"].Style.Font.Name = "Calibri";
                plist2.Cells["H13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["H13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["H13"].Value = "PESO NETO";


                plist2.Cells["I13"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["I13"].Style.Font.Size = 11;
                plist2.Cells["I13"].Style.Font.Bold = true;
                plist2.Cells["I13"].Style.Font.Name = "Calibri";
                plist2.Cells["I13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["I13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["I13"].Value = "PESO BRUTO";
                // fin de 4 headers


                plist2.Cells["A14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["A14"].Style.Font.Size = 10;
                plist2.Cells["A14"].Style.Font.Name = "Calibri";
                plist2.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A14"].Style.WrapText = true;
                plist2.Cells["A14"].Value = "FORMATO";


                plist2.Cells["B14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["B14"].Style.Font.Size = 10;
                plist2.Cells["B14"].Style.Font.Name = "Calibri";
                plist2.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["B14"].Style.WrapText = true;
                plist2.Cells["B14"].Value = "MODELO";
                /*
                plist2.Cells["B14:B15"].Merge = true;
                plist2.Cells["B14:B15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["B14"].Style.Font.Size = 10;
                plist2.Cells["B14"].Style.Font.Name = "Calibri";
                plist2.Cells["B14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["B14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["B14"].Style.WrapText = true;
                plist2.Cells["B14"].Value = "MODEL";
                */

                plist2.Cells["C14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["C14"].Style.Font.Size = 10;
                plist2.Cells["C14"].Style.Font.Name = "Calibri";
                plist2.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["C14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["C14"].Style.WrapText = true;
                plist2.Cells["C14"].Value = "CLASE";


                plist2.Cells["D14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["D14"].Style.Font.Size = 10;
                plist2.Cells["D14"].Style.Font.Name = "Calibri";
                plist2.Cells["D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["D14"].Style.WrapText = true;
                plist2.Cells["D14"].Value = "PALLETS";


                plist2.Cells["E14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["E14"].Style.Font.Size = 10;
                plist2.Cells["E14"].Style.Font.Name = "Calibri";
                plist2.Cells["E14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["E14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["E14"].Style.WrapText = true;
                plist2.Cells["E14"].Value = "CAJAS";

                plist2.Cells["F14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["F14"].Style.Font.Size = 10;
                plist2.Cells["F14"].Style.Font.Name = "Calibri";
                plist2.Cells["F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["F14"].Style.WrapText = true;
                plist2.Cells["F14"].Value = "M²";


                plist2.Cells["G14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["G14"].Style.Font.Size = 10;
                plist2.Cells["G14"].Style.Font.Name = "Calibri";
                plist2.Cells["G14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["G14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["G14"].Style.WrapText = true;
                plist2.Cells["G14"].Value = "PIEZAS";
            
                plist2.Cells["H14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["H14"].Style.Font.Size = 10;
                plist2.Cells["H14"].Style.Font.Name = "Calibri";
                plist2.Cells["H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["H14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["H14"].Style.WrapText = true;
                plist2.Cells["H14"].Value = "PESO";

                plist2.Cells["I14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["I14"].Style.Font.Size = 10;
                plist2.Cells["I14"].Style.Font.Name = "Calibri";
                plist2.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["I14"].Style.WrapText = true;
                plist2.Cells["I14"].Value = "PESO";
                //plist2.Cells["I13"].Value = "" + currency;
                /*
                plist2.Cells["I14:I15"].Merge = true;
                plist2.Cells["I14:I15"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["I14"].Style.Font.Size = 10;
                plist2.Cells["I14"].Style.Font.Name = "Calibri";
                plist2.Cells["I14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["I14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["I14"].Style.WrapText = true;
                plist2.Cells["I14"].Value = "No.of CTNRS";
                */


                #endregion headerpackinglistv2

                #region builerpackinglistv2
                rownu = 15; contu = 1;
                caja = ""; mtscaja = ""; kgspiece = ""; kgscaja = ""; sku = ""; 

                ctns = 0; sqm = 0; tctns = 0; tsqm = 0; cajad = 0; mtsd = 0; pallets = 0;
                double kilosnetos = 0 ,kilosbrutos=0;
                tonelaje = 0;
                size = ""; sizel = ""; pallet = ""; units = "";
                //ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                //ord, cant, clave, ume, pu, importe, container
                string querycpl = "SELECT container,precinto FROM containersipl WHERE ord ='" + id + "' ORDER BY id ASC;";
                SqlCommand cmcpl = new SqlCommand(querycpl, con);
                SqlDataAdapter dacpl = new SqlDataAdapter(cmcpl);
                DataTable dtcpl = new DataTable();
                dacpl.Fill(dtcpl);
                int cuentacpl = dtcpl.Rows.Count;
                if (cuentacpl > 0)
                {
                    foreach (DataRow rowc in dtcpl.Rows)
                    {


                        string querycp = "SELECT * FROM rowsipl WHERE ord ='" + id + "' AND container='" + rowc["container"] + "' ORDER BY id ASC;";
                        SqlCommand cmcp = new SqlCommand(querycp, con);
                        SqlDataAdapter dacp = new SqlDataAdapter(cmcp);
                        DataTable dtcp = new DataTable();
                        dacp.Fill(dtcp);
                        int cuentacp = dtcp.Rows.Count;
                        if (cuentacp > 0)
                        {

                            foreach (DataRow row in dtcp.Rows)
                            {

                                string querya = "SELECT * FROM artsipl WHERE clave ='" + row["clave"] + "';";
                                SqlCommand cma = new SqlCommand(querya, con);
                                SqlDataAdapter daa = new SqlDataAdapter(cma);
                                DataTable dta = new DataTable();
                                daa.Fill(dta);
                                int cuentaa = dta.Rows.Count;
                                if (cuentaa > 0)
                                {
                                    foreach (DataRow rowa in dta.Rows)
                                    {
                                        size = "" + rowa["size"]; sizel = "" + rowa["sizel"];
                                        caja = "" + rowa["caja"]; mtscaja = "" + rowa["mtscaja"]; kgspiece = "" + rowa["kgspiece"];
                                        kgscaja = "" + rowa["kgscaja"]; sku = "" + rowa["size"];
                                        pallet = "" + rowa["pallet"];
                                        units = "" + rowa["ume"];
                                       
                                    }

                                }
                                cma.Dispose(); daa.Dispose(); dta.Dispose();


                                //ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                                //ord, cant, clave, ume, pu, importe, container
                                plist2.Cells["A" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["A" + rownu].Style.Font.Size = 9;
                                plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["A" + rownu].Style.WrapText = true;
                                plist2.Cells["A" + rownu].Value = "" + size + "X" + sizel ;

                                plist2.Cells["B" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["B" + rownu].Style.Font.Size = 9;
                                plist2.Cells["B" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["B" + rownu].Style.WrapText = true;
                                plist2.Cells["B" + rownu].Value = "" + row["clave"];

                                plist2.Cells["C" + rownu].Merge = true;
                                plist2.Cells["C" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["C" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["C" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["C" + rownu].Style.Font.Size = 9;
                                plist2.Cells["C" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["C" + rownu].Style.WrapText = true;
                                plist2.Cells["C" + rownu].Value = 1;

                                /*
                                plist2.Cells["B" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["B" + rownu].Style.Font.Size = 9;
                                plist2.Cells["B" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["B" + rownu].Style.WrapText = true;
                                plist2.Cells["B" + rownu].Value ="" + row["clave"];
                                */


                                //plist2.Cells["C" + rownu].Value = "" + units; (M2 OR PIECE)


                                sqm = 0;
                                tctns = 0;
                                tsqm = 0;
                                cajad = 0;
                                mtsd = 0;
                                pallets = 0;
                                pieces = 0;

                                try { ctns = double.Parse("" + caja); }
                                catch { ctns = 0; }

                                try { sqm = double.Parse("" + mtscaja); }
                                catch { sqm = 0; }

                                try { pallets = double.Parse("" + pallet); }
                                catch { pallets = 0; }

                                try
                                {
                                    tctns = double.Parse("" + row["cant"]) / ctns;
                                }
                                catch { tctns = 0; }

                                try
                                {
                                    tsqm = sqm * tctns;
                                }
                                catch { tsqm = 0; }

                                try
                                {
                                    pieces = (ctns * double.Parse("" + row["cant"])) / sqm;
                                }
                                catch
                                {
                                    pieces = 0;
                                }

                                if (row["pallets"].ToString() != "")
                                {
                                    pallets = double.Parse("" + row["pallets"]);
                                }
                                else
                                {
                                    try
                                    {
                                        pallets = double.Parse("" + row["cant"]) / pallets;
                                    }
                                    catch { pallets = 1; }
                                    //if (pallets < 1) { pallets = 1; }
                                    if (pallets.ToString() == "∞") { pallets = 0; }
                                }
                                try
                                {
                                    cajad = double.Parse("" + row["cant"]) / sqm;
                                }
                                catch { cajad = 0; }

                                kilosnetos = 0;  kilosbrutos = 0;

                                kilosnetos = (cajad * double.Parse ("" + kgscaja));
                                kilosbrutos = (pallets * double.Parse("" + palletkgs.Text.ToString())) + kilosnetos;


                                plist2.Cells["D" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["D" + rownu].Style.Font.Size = 9;
                                plist2.Cells["D" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["D" + rownu].Style.WrapText = true;
                                plist2.Cells["D" + rownu].Style.Numberformat.Format = "#,##0.00";
                                plist2.Cells["D" + rownu].Value = pallets;
                                //plist2.Cells["D" + rownu].Value = double.Parse("" + row["cant"]);

                                plist2.Cells["E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["E" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["E" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["E" + rownu].Style.Font.Size = 9;
                                plist2.Cells["E" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["E" + rownu].Style.WrapText = true;
                                plist2.Cells["E" + rownu].Style.Numberformat.Format = "#,##0";
                                plist2.Cells["E" + rownu].Value = double.Parse("" + cajad);

                                plist2.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["F" + rownu].Style.Font.Size = 9;
                                plist2.Cells["F" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["F" + rownu].Style.WrapText = true;
                                plist2.Cells["F" + rownu].Style.Numberformat.Format = "#,##0.00";
                                plist2.Cells["F" + rownu].Value = double.Parse("" + row["cant"]);

                                plist2.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["G" + rownu].Style.Font.Size = 9;
                                plist2.Cells["G" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["G" + rownu].Style.WrapText = true;
                                plist2.Cells["G" + rownu].Style.Numberformat.Format = "#,##0";
                                plist2.Cells["G" + rownu].Value = pieces;
                                //plist2.Cells["G" + rownu].Value = double.Parse("" + row["pu"]);

                                
                                plist2.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["H" + rownu].Style.Font.Size = 9;
                                plist2.Cells["H" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["H" + rownu].Style.WrapText = true;
                                plist2.Cells["H" + rownu].Style.Numberformat.Format = "#,##0.00";
                                plist2.Cells["H" + rownu].Value = double.Parse("" + row["pesoneto"]);


                                plist2.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                plist2.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                plist2.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                plist2.Cells["I" + rownu].Style.Font.Size = 9;
                                plist2.Cells["I" + rownu].Style.Font.Name = "Calibri";
                                plist2.Cells["I" + rownu].Style.WrapText = true;
                                plist2.Cells["I" + rownu].Style.Numberformat.Format = "#,##0.00";
                                plist2.Cells["I" + rownu].Value = double.Parse("" + row["pesobruto"]);



                                rownu = rownu + 1; contu = contu + 1;



                               
                            }
                        }
                        
                        plist2.Cells["A" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        plist2.Cells["A" + rownu].Style.Font.Size = 8;
                        plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                        plist2.Cells["A" + rownu].Style.WrapText = true;
                        plist2.Cells["A" + rownu].Value = "CONTENEDOR:";

                        plist2.Cells["B" + rownu + ":C" + rownu].Merge = true;
                        plist2.Cells["B" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        plist2.Cells["C" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        plist2.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        plist2.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        plist2.Cells["B" + rownu].Style.Font.Size = 8;
                        plist2.Cells["B" + rownu].Style.Font.Name = "Calibri";
                        plist2.Cells["B" + rownu].Style.WrapText = true;
                        plist2.Cells["B" + rownu].Value = "" + rowc["container"];

                        plist2.Cells["D" + rownu + ":G" + rownu].Merge = true;
                        plist2.Cells["D" + rownu + ":G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        plist2.Cells["D" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        plist2.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        plist2.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        plist2.Cells["D" + rownu].Style.Font.Size = 8;
                        plist2.Cells["D" + rownu].Style.Font.Name = "Calibri";
                        plist2.Cells["D" + rownu].Style.WrapText = true;
                        plist2.Cells["D" + rownu].Value = "PRECINTO: " + rowc["precinto"];

                       




                        plist2.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        plist2.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        plist2.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        plist2.Cells["H" + rownu].Style.Font.Size = 8;
                        plist2.Cells["H" + rownu].Style.Font.Name = "Calibri";
                        plist2.Cells["H" + rownu].Style.WrapText = true;
                        plist2.Cells["H" + rownu].Style.Numberformat.Format = "#,##0.00";
                        plist2.Cells["H" + rownu].Formula = "=SUM(H" + (rownu - cuentacp) + ":H" + (rownu - 1) + ")";
                        plist2.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        
                        plist2.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        plist2.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        plist2.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        plist2.Cells["I" + rownu].Style.Font.Size = 8;
                        plist2.Cells["I" + rownu].Style.Font.Name = "Calibri";
                        plist2.Cells["I" + rownu].Style.WrapText = true;
                        plist2.Cells["I" + rownu].Style.Numberformat.Format = "#,##0.00";
                        plist2.Cells["I" + rownu].Formula = "=SUM(I" + (rownu - cuentacp) + ":I" + (rownu - 1) + ")";
                        plist2.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);



                        rownu = rownu + 1; contu = contu + 1;


                    }
                }
                cmcpl.Dispose(); dacpl.Dispose(); dtcpl.Dispose();
                cm.Dispose(); da.Dispose(); dt.Dispose();
                #endregion builerpackinglistv2

                #region footerpackinglistv2


                rowf = rownu;
                if (rowf > rowni)
                {
                    rowf = rowf - 1;
                }
                rownu = rownu + 1;

                plist2.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["G" + rownu].Style.Font.Size = 9;
                plist2.Cells["G" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["G" + rownu].Style.WrapText = true;
                plist2.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                plist2.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["H" + rownu].Style.Font.Size = 9;
                plist2.Cells["H" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["H" + rownu].Style.WrapText = true;
                plist2.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["H" + rownu].Style.Numberformat.Format = "" + csymbol;
                plist2.Cells["H" + rownu].Formula = "=SUM(H" + rowni + ":H" + rowf + ")";





                plist2.Cells["D" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["D" + rownu].Style.Font.Size = 9;
                plist2.Cells["D" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["D" + rownu].Style.WrapText = true;
                plist2.Cells["D" + rownu].Value = "PALLETS";


                plist2.Cells["E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["E" + rownu].Style.Font.Size = 9;
                plist2.Cells["E" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["E" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["E" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["E" + rownu].Style.WrapText = true;
                plist2.Cells["E" + rownu].Value = "CAJAS";

                plist2.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["F" + rownu].Style.Font.Size = 9;
                plist2.Cells["F" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["F" + rownu].Style.WrapText = true;
                plist2.Cells["F" + rownu].Value = "M²";


                plist2.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["G" + rownu].Style.Font.Size = 9;
                plist2.Cells["G" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["G" + rownu].Style.WrapText = true;
                plist2.Cells["G" + rownu].Value = "PIEZAS";

                plist2.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["H" + rownu].Style.Font.Size = 9;
                plist2.Cells["H" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["H" + rownu].Style.WrapText = true;
                plist2.Cells["H" + rownu].Value = "PESO NETO";

                plist2.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["I" + rownu].Style.Font.Size = 9;
                plist2.Cells["I" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["I" + rownu].Style.WrapText = true;
                plist2.Cells["I" + rownu].Value = "PESO BRUTO";


                rownu = rownu + 1;



                plist2.Cells["D" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["D" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["D" + rownu].Style.Font.Size = 9;
                plist2.Cells["D" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["D" + rownu].Style.WrapText = true;
                plist2.Cells["D" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["D" + rownu].Style.Numberformat.Format = "##0.00";
                plist2.Cells["D" + rownu].Formula = "=SUM(D" + rowni + ":D" + (rowf -1)+ ")";


                plist2.Cells["E" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["E" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["E" + rownu].Style.Font.Size = 9;
                plist2.Cells["E" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["E" + rownu].Style.WrapText = true;
                plist2.Cells["E" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["E" + rownu].Style.Numberformat.Format = "##0";
                plist2.Cells["E" + rownu].Formula = "=SUM(E" + rowni + ":E" + (rowf - 1) + ")";


                plist2.Cells["G" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["G" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["G" + rownu].Style.Font.Size = 9;
                plist2.Cells["G" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["G" + rownu].Style.WrapText = true;
                plist2.Cells["G" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["G" + rownu].Style.Numberformat.Format = "##0";
                plist2.Cells["G" + rownu].Formula = "=SUM(G" + rowni + ":G" + (rowf - 1) + ")";



                plist2.Cells["F" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["F" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["F" + rownu].Style.Font.Size = 9;
                plist2.Cells["F" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["F" + rownu].Style.WrapText = true;
                plist2.Cells["F" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["F" + rownu].Style.Numberformat.Format = "##0.00";
                plist2.Cells["F" + rownu].Formula = "=SUM(F" + rowni + ":F" + (rowf - 1) + ")";

                plist2.Cells["H" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["H" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["H" + rownu].Style.Font.Size = 9;
                plist2.Cells["H" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["H" + rownu].Style.WrapText = true;
                plist2.Cells["H" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["H" + rownu].Style.Numberformat.Format = "#,##0.00";
                plist2.Cells["H" + rownu].Value = double.Parse("" + netofact);
               

                plist2.Cells["I" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                plist2.Cells["I" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["I" + rownu].Style.Font.Size = 9;
                plist2.Cells["I" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["I" + rownu].Style.WrapText = true;
                plist2.Cells["I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["I" + rownu].Style.Numberformat.Format = "#,##0.00";
                plist2.Cells["I" + rownu].Value = double.Parse("" + brutosfact);
                

                rownu = rownu + 2;


                plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["A" + rownu].Style.Font.Size = 9;
                plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["A" + rownu].Style.WrapText = true;
                plist2.Cells["A" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["A" + rownu].Value = "DESTINO:";

                plist2.Cells["B" + rownu + ":I" + rownu].Merge = true;
                plist2.Cells["B" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                plist2.Cells["B" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                plist2.Cells["B" + rownu].Style.Font.Size = 9;
                plist2.Cells["B" + rownu].Style.Font.Name = "Calibri";
                plist2.Cells["B" + rownu].Style.WrapText = true;
                plist2.Cells["B" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                plist2.Cells["B" + rownu].Value = "" + origdest;

                rownu = rownu + 1;

                
                                if (obs1 != "")
                                {
                                    int rowes = 32;
                                    rownu = rownu + 2;
                                    plist2.Cells["A" + rownu + ":I" + rownu].Merge = true;
                                    plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    plist2.Cells["A" + rownu].Style.Font.Size = 9;
                                    plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                                    plist2.Cells["A" + rownu].Style.Font.Bold = true;
                                    plist2.Cells["A" + rownu].Style.WrapText = true;
                                    plist2.Row(rownu).Height = 15;
                                    plist2.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    plist2.Cells["A" + rownu].Value = "OBSERVACIONES:";

                                    rownu = rownu + 1;
                                    plist2.Cells["A" + rownu + ":I" + rownu].Merge = true;
                                    plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    plist2.Cells["A" + rownu].Style.Font.Size = 9;
                                    plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                                    plist2.Cells["A" + rownu].Style.WrapText = true;
                                    plist2.Row(rownu).Height = rowes;
                                    plist2.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    plist2.Cells["A" + rownu].Value = "" + obs1;

                                    if (obs2 != "")
                                    {
                                        rownu = rownu + 1;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Merge = true;
                                        plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.Font.Size = 9;
                                        plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                                        plist2.Cells["A" + rownu].Style.WrapText = true;
                                        plist2.Row(rownu).Height = rowes;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        plist2.Cells["A" + rownu].Value = "" + obs2;
                                    }

                                    if (obs3 != "")
                                    {
                                        rownu = rownu + 1;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Merge = true;
                                        plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.Font.Size = 9;
                                        plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                                        plist2.Cells["A" + rownu].Style.WrapText = true;
                                        plist2.Row(rownu).Height = rowes;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        plist2.Cells["A" + rownu].Value = "" + obs3;
                                    }

                                    if (obs4 != "")
                                    {
                                        rownu = rownu + 1;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Merge = true;
                                        plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.Font.Size = 9;
                                        plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                                        plist2.Cells["A" + rownu].Style.WrapText = true;
                                        plist2.Row(rownu).Height = rowes;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        plist2.Cells["A" + rownu].Value = "" + obs4;
                                    }
                                    if (obs5 != "")
                                    {
                                        rownu = rownu + 1;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Merge = true;
                                        plist2.Cells["A" + rownu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        plist2.Cells["A" + rownu].Style.Font.Size = 9;
                                        plist2.Cells["A" + rownu].Style.Font.Name = "Calibri";
                                        plist2.Cells["A" + rownu].Style.WrapText = true;
                                        plist2.Row(rownu).Height = rowes;
                                        plist2.Cells["A" + rownu + ":I" + rownu].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                        plist2.Cells["A" + rownu].Value = "" + obs5;
                                    }
                                }
                
                netofull = 0; brutofull = 0;

                #endregion footerpackinglistv2

                #region opener

                proforma.Column(1).AutoFit();
                proforma.Column(1).Width = 10;

                proforma.Column(2).AutoFit();
                proforma.Column(2).Width = 20;
                proforma.Column(3).AutoFit();
                proforma.Column(3).Width = 6;
                proforma.Column(4).Width = 8;
                proforma.Column(5).Width = 8;
                proforma.Column(6).Width = 8;
                proforma.Column(7).AutoFit();
                proforma.Column(7).Width = 9;
                proforma.Column(8).Width = 13;
                proforma.Column(9).Width = 15;
                //proforma.Column(8).AutoFit();

                plist2.Column(1).AutoFit();
                plist2.Column(1).Width = 10;

                plist2.Column(2).AutoFit();
                plist2.Column(2).Width = 20;
                plist2.Column(3).AutoFit();
                plist2.Column(3).Width = 6;
                plist2.Column(4).Width = 8;
                plist2.Column(5).Width = 8;
                plist2.Column(6).Width = 8;
                plist2.Column(7).AutoFit();
                plist2.Column(7).Width = 9;
                plist2.Column(8).Width = 13;
                plist2.Column(9).Width = 15;

                double TOPMA = 0;
                double LEFTMA = 0;


                try {

                    TOPMA = double.Parse(supo.Text);
                }
                catch { TOPMA = 10; }
                try {
                    LEFTMA = double.Parse(lefte.Text);
                }
                catch { LEFTMA = 10; }

                try { TOPMA = TOPMA / 10; }
                catch { TOPMA = 1; }

                try { LEFTMA = LEFTMA / 10; }
                catch { LEFTMA = 1; }

                proforma.PrinterSettings.TopMargin = (decimal)TOPMA / 2.54M; // narrow border
                proforma.PrinterSettings.RightMargin = (decimal).4 / 2.54M; //narrow border
                proforma.PrinterSettings.LeftMargin = (decimal)LEFTMA / 2.54M; //narrow border
                proforma.PrinterSettings.BottomMargin = (decimal).4 / 2.54M; //narrow border

                plist2.PrinterSettings.TopMargin = (decimal)TOPMA / 2.54M; // narrow border
                plist2.PrinterSettings.RightMargin = (decimal).4 / 2.54M; //narrow border
                plist2.PrinterSettings.LeftMargin = (decimal)LEFTMA / 2.54M; //narrow border
                plist2.PrinterSettings.BottomMargin = (decimal).4 / 2.54M; //narrow border

                //proforma.Row(30).PageBreak = true;
                proforma.PrinterSettings.PaperSize = ePaperSize.Letter;
                proforma.PrinterSettings.Orientation = eOrientation.Portrait;
                //proforma.PrinterSettings.Scale = 75;


                //concentrado.Row(30).PageBreak = true;
                plist2.PrinterSettings.PaperSize = ePaperSize.Letter;
                plist2.PrinterSettings.Orientation = eOrientation.Portrait;
                //concentrado.PrinterSettings.Scale = 75;

                excel.SaveAs(excelFile);
                System.Diagnostics.Process.Start(@"" + excelFile);
                con.Close();
                #endregion opener
            }
            else { MessageBox.Show("Debes seleccionar un Invoice"); }

        }

        private void list_Load(object sender, EventArgs e)
        {
            deci = Int32.Parse("" + deces.Value);
            string inicial="", final="";
            DateTime Hoy = DateTime.Today;
            DateTime dt = Hoy;
            dt = dt.AddDays(-(dt.Day - 1));
            init.Text = "" + dt.ToString("dd-MM-yyyy");
            ejer.Text = "" + dt.ToString("yyyy");
            DateTime dtTo = dt;
            dtTo = dtTo.AddMonths(1);
            dtTo = dtTo.AddDays(-(dtTo.Day));
            finit.Text = "" + dtTo.ToString("dd-MM-yyyy");

            SqlConnection con = new
            SqlConnection("" + config.cade);
            con.Open();

            inicial = "" + init.Text; final = "" + finit.Text;

            var primera = Convert.ToDateTime(inicial);
            var segunda = Convert.ToDateTime(final);

            inicial  ="" + primera.ToString("yyyy-MM-dd");
            final = "" + segunda.ToString("yyyy-MM-dd");
            string range = "";
            range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
            SqlDataAdapter DA = new SqlDataAdapter();
            string sqlSelectAll = invoice_query + " " + range + ";";
            DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

            DataTable table = new DataTable();
            DA.Fill(table);

            BindingSource bSource = new BindingSource();
            bSource.DataSource = table;
            lister.DataSource = bSource;
            supo.Text = "" + Properties.Settings.Default.topmargin;
            lefte.Text = "" + Properties.Settings.Default.leftmargin;
            resizegrid();


            string query = "SELECT TOP(1)* FROM configs ORDER BY id ASC;";
            SqlCommand cmc = new SqlCommand(query, con);
            SqlDataAdapter dac = new SqlDataAdapter(cmc);
            DataTable dtc = new DataTable();
            dac.Fill(dtc);
            int cuentac = dtc.Rows.Count;
            if (cuentac > 0)
            {
                foreach (DataRow row in dtc.Rows)
                {
                    palletkgs.Text = "" + row["tarima"];
                }
            }
            dtc.Dispose(); cmc.Dispose(); dac.Dispose();


        }

        private void lister_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var dataIndexNo = lister.Rows[e.RowIndex].Index.ToString();
                string cellValue = lister.Rows[e.RowIndex].Cells[11].Value.ToString();

                //MessageBox.Show("The row index = " + dataIndexNo.ToString() + " and the row data in second column is: "
                //    + cellValue.ToString());

                Form editor = new editor();
                config.idinvoice = "" + cellValue;
                editor.ShowDialog(this);


                iplcont.DataSource = null;
                BindingSource bSource = new BindingSource();
                bSource.DataSource = null;
                string inicial = "", final = "", cli = "", emi = "";
                SqlConnection con = new SqlConnection("" + config.cade);
                con.Open();
                inicial = "" + init.Text; final = "" + finit.Text;

                var primera = Convert.ToDateTime(inicial);
                var segunda = Convert.ToDateTime(final);
                inicial = "" + primera.ToString("yyyy-MM-dd");
                final = "" + segunda.ToString("yyyy-MM-dd");
                string range = "";

                cli = "" + numc.Text;
                emi = "" + nume.Text;

                if (cli != "") { cli = " AND  idcli='" + cli + "' "; }
                if (emi != "") { emi = " AND  empresa='" + emi + "' "; }
                range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE) " + cli + " " + emi + ";";
                SqlDataAdapter DA = new SqlDataAdapter();
                string sqlSelectAll = invoice_query + " " + range + ";";
                DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                DataTable table = new DataTable();
                DA.Fill(table);
                bSource.DataSource = table;
                lister.DataSource = bSource;
                resizegrid();

            }
            catch
            { }
        }
        private void lister_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                iplcont.DataSource = null;
                selecto.Text = "";
                var dataIndexNo = lister.Rows[e.RowIndex].Index.ToString();
                string cellValue = lister.Rows[e.RowIndex].Cells[0].Value.ToString();
                string cellValuen = lister.Rows[e.RowIndex].Cells[1].Value.ToString();
                string cellValueid = lister.Rows[e.RowIndex].Cells[11].Value.ToString();

                selecto.Text = "" + cellValue.ToString();
                idinvo.Text = "" + cellValueid.ToString();
                numee.Text = "" + cellValuen.ToString();
                SqlConnection con = new
                SqlConnection("" + config.cade);
                con.Open();
                SqlDataAdapter DA = new SqlDataAdapter();
                string sqlSelectAll = "Select cant AS Cant, clave AS Clave, FORMAT( convert(numeric(18,5),replace(pu,',','')),'###,###,###.00000','ES-mx')  AS PU, FORMAT( convert(numeric(18,5),replace( importe,',','')),'###,###,###.00000','ES-mx') AS Monto, container AS Cont FROM rowsipl WHERE ord='" + cellValueid + "' ORDER BY id ASC;";
                DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                DataTable table = new DataTable();
                DA.Fill(table);

                BindingSource bSource = new BindingSource();
                bSource.DataSource = table;
                iplcont.DataSource = bSource;
                DA.Dispose();
                resizegriddet();
                con.Close();
            }
            else
            {
                selecto.Text = "";idinvo.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            iplcont.DataSource = null;
            BindingSource bSource = new BindingSource();
            bSource.DataSource = null;
            string inicial = "", final = "",cli="",emi="";
            SqlConnection con = new SqlConnection("" + config.cade);
            con.Open();
            inicial = "" + init.Text; final = "" + finit.Text;

            var primera = Convert.ToDateTime(inicial);
            var segunda = Convert.ToDateTime(final);
            inicial = "" + primera.ToString("yyyy-MM-dd");
            final = "" + segunda.ToString("yyyy-MM-dd");
            string range = "";

            cli = "" + numc.Text;
            emi = "" + nume.Text;

            if (cli != "") { cli = " AND  idcli='" + cli + "' "; }
            if (emi != "") { emi = " AND  empresa='" +  emi + "' ";   }
            range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE) " + cli + " " + emi + ";";
            SqlDataAdapter DA = new SqlDataAdapter();
            string sqlSelectAll =  invoice_query + " "+ range  + ";";
            DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

            DataTable table = new DataTable();
            DA.Fill(table);
            bSource.DataSource = table;
            lister.DataSource = bSource;
            resizegrid();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) )
            {
                e.Handled = true;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        

        private void button2_Click(object sender, EventArgs e)
        {

            if (idinvo.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea CANCELAR?\r\nInvoice/packing list #: " + selecto.Text + ".", "--Cancelar Set de Documentos--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    string cellValueid = idinvo.Text;
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "UPDATE  invoicespl SET stats='CANCELADA' WHERE id=" + cellValueid;
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    /*
                    qu = "DELETE FROM containersipl WHERE ord='" + cellValueid + "';";
                    myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    qu = "DELETE FROM invoicespl WHERE id=" + cellValueid;
                    myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    qu = "DELETE FROM rowsservpl WHERE ord='" + config.idinvoice + "';";
                    myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();
                    */

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = null;
                    string inicial = "", final = "", cli = "", emi = "";
                    inicial = "" + init.Text; final = "" + finit.Text;

                    var primera = Convert.ToDateTime(inicial);
                    var segunda = Convert.ToDateTime(final);
                    inicial = "" + primera.ToString("yyyy-MM-dd");
                    final = "" + segunda.ToString("yyyy-MM-dd");
                    string range = "";

                    cli = "" + numc.Text;
                    emi = "" + nume.Text;

                    if (cli != "") { cli = " AND  " + cli; }
                    if (emi != "") { }
                    range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE) ";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    string sqlSelectAll = invoice_query + " " + range + ";";
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);
                    bSource.DataSource = table;
                    lister.DataSource = bSource;
                    iplcont.DataSource = null;
                    resizegrid();

                }
            }
            else { MessageBox.Show("Debes seleccionar un Invoice para cancelarlo."); }

        }

        private void resizegrid()
        {
            lister.Columns[0].Width = 80;
            lister.Columns[1].Width = 60;
            lister.Columns[2].Width = 30;
            lister.Columns[3].Width = 30;
            lister.Columns[4].Width = 110;
            lister.Columns[5].Width = 75;
            lister.Columns[6].Width = 55;
            lister.Columns[7].Width = 75;
            lister.Columns[8].Width = 50;
            lister.Columns[9].Width = 75;
            lister.Columns[10].Width = 75;
            lister.Columns[11].Visible = false;
            lister.Columns[12].Width = 75;
            selecto.Text = ""; idinvo.Text = ""; numee.Text = "";
        }
        private void resizegriddet()
        {
            iplcont.Columns[0].Width = 65;
            iplcont.Columns[1].Width = 120;
            iplcont.Columns[2].Width = 60;
            iplcont.Columns[3].Width = 88;
            iplcont.Columns[4].Width = 120;
        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            /*
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
            */
        }

        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            /*
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
            */
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.topmargin = supo.Text;
            Properties.Settings.Default.leftmargin = lefte.Text;
            Properties.Settings.Default.Save();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (idinvo.Text == "") { MessageBox.Show("No selecciono ningun documento para imprimir"); }
            else
            {
             #region init            
                DateTime hoy = DateTime.Now;
                logpesos.Text = "";
                string dates = hoy.Day.ToString() + "-" + hoy.Month.ToString() + "-" + hoy.Year.ToString() + "-" + hoy.Hour.ToString() + "-" + hoy.Minute.ToString() + "-" + hoy.Second.ToString();
                string paths = @"" + config.tempofiles + @"\INVOICE_PACKINGLIST_" + dates + ".pdf";
                // Create a new PDF document
                PdfDocument document = new PdfDocument();


                document.Info.Title = "Factura y packing list número: " + numee.Text;
                // Create an empty page Y PONERLE SUS PROPIEDADES
                PdfPage page = document.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Portrait;
                page.Size = PdfSharp.PageSize.Letter;
                // page = document.AddPage();
                // Get an XGraphics object for drawing
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XTextFormatter tf = new XTextFormatter(gfx);
                // FUENTES PARA EL DOCTO
                XFont dieznegra = new XFont("Arial", 10, XFontStyle.Bold);
                XFont catonegra = new XFont("Arial", 14, XFontStyle.Bold);

                XFont ocho = new XFont("Arial", 8, XFontStyle.Regular);
                XFont DOCE = new XFont("Arial", 12, XFontStyle.Regular);
                XFont ochoneg = new XFont("Arial", 8, XFontStyle.Bold);
                XFont SIETE = new XFont("Arial", 7, XFontStyle.Regular);
                XFont SEIS = new XFont("Arial", 6, XFontStyle.Regular);
                XFont CINCO = new XFont("Arial", 5, XFontStyle.Regular);
                XPen pen = new XPen(XColors.Black, 1);
                XPen peng = new XPen(XColors.Gray, 0.3);
                string id = "", numees = "", folio = "", empresa = "", idcli = "", nomcli = "", callecli = "", numcli = "", numclii = "", colcli = "", muncli = "", edocli = "", paiscli = "", fecha = "", albaran = "", origdest = "", tot = "", currency = "", cpcli = "", nifcli = "";
                string ide = "", nome = "", nife = "", callee = "", nume = "", numie = "", cole = "", cde = "", estadoe = "", paise = "", cpe = "";
                string netofact = "", brutosfact = "";
                string obs1 = "", obs2 = "", obs3 = "", obs4 = "", obs5 = "";
                string csymbol_l = "";
                string csymbol_r = "";

                string size = "", sizel = "";
                string caja = "", mtscaja = "", kgspiece = "";
                string kgscaja = "", sku = "";
                string pallet = "";
                string units = "";


                double kgscontainer = 0, kgsacums = 0;
                int cuantasvan = 0, cuantastotoal = 0;

                double mtspartidas = 0, mtsacum = 0, kgs1caja = 0, mts1caja = 0, tonelaje = 0, tonelajecontainer = 0;
                double tarimaskgs = 0, kgstar = 0, kilosXpallets = 0;
                double netofull = 0, brutofull = 0;
                double cajastot = 0, palletstot = 0, piezastot = 0;
                string cellrow1 = "", cellrow2 = "";
                int inity = 70, y = 25, x = 100;

                double sqm = 0;
                double tctns = 0;
                double tsqm = 0;
                double cajad = 0;
                double mtsd = 0;
                double pallets = 0;
                double pieces = 0;
                double ctns = 0;
                int totalpartidas = 0;

                #endregion init

             #region headerproform

                //concentrado.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //concentrado.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


                SqlConnection con = new SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT * FROM invoicespl WHERE id =" + idinvo.Text + ";";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {

                        id = "" + row["id"]; folio = "" + row["folio"]; empresa = "" + row["empresa"]; idcli = "" + row["idcli"];
                        nomcli = "" + row["nomcli"];
                        callecli = "" + row["callecli"];
                        numcli = "" + row["numcli"]; numclii = "" + row["numclii"]; colcli = "" + row["colcli"];
                        muncli = "" + row["muncli"];
                        edocli = "" + row["edocli"]; paiscli = "" + row["paiscli"];
                        fecha = "" + row["fecha"];
                        origdest = "" + row["origdest"]; tot = "" + row["tot"];
                        currency = "" + row["currency"];

                        obs1 = "" + row["obs1"];
                        obs2 = "" + row["obs2"];
                        obs3 = "" + row["obs3"];
                        obs4 = "" + row["obs4"];
                        obs5 = "" + row["obs5"];

                        albaran = "" + row["albaran"];
                        numees = "" + row["number"];
                        netofact = "" + row["pesoneto"]; brutosfact = "" + row["pesobruto"];
                    }
                }
                else
                {
                    MessageBox.Show("Orden no existe");
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT * FROM empresasipl WHERE id =" + empresa + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT id,cp,nif FROM clientesipl WHERE id =" + idcli + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        cpcli = "" + row["cp"];
                        nifcli = "" + row["nif"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                if (currency == "USD") { csymbol_l = "US $"; }
                if (currency == "EUR") { csymbol_r = " €"; }




                query = "SELECT id FROM rowsipl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    totalpartidas = totalpartidas + cuenta;
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                query = "SELECT id  FROM rowsservpl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    totalpartidas = totalpartidas + cuenta;
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                //MessageBox.Show(" total de partidas para pintar: " + totalpartidas);
                cuenta = 0;

                //creacion de las listas para productos y servicios y saber cuales estan inmpresos y cuales no
                List<string> rowids = new List<string>();
                List<string> rowidserv = new List<string>();
                int conteo = 0;
                int contadoarr = 0;

                query = "SELECT id FROM rowsipl WHERE ord ='" + idinvo.Text + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        rowids.Add(row["id"] + "|no");
                        conteo = conteo + 1;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT id FROM rowsservpl WHERE ord ='" + idinvo.Text + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        rowidserv.Add(row["id"] + "|no");
                        conteo = conteo + 1;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();




                //arrfilas[conteo] = "no|" + row["id"];
                String[] arrfilas = rowids.ToArray();
                String[] arrfilasedit = rowids.ToArray();
                //FIN DE LA CREACION creacion de las listas para productos y servicios y saber cuales estan inmpresos y cuales no


                //FOR EACH MAESTRO DE PIEZAS Y SERVICIOS

                cajastot =0;
                palletstot = 0;
                piezastot = 0;
                for (int i = 1; i <= totalpartidas; i++)
                {
                    /*
                    string logo = @"" + Path.GetDirectoryName(Application.ExecutablePath) + @"\" + "logo.jpg";
                    if (!File.Exists(@"" + logo))
                    {
                        //throw new FileNotFoundException(String.Format("No se encuentra el Logo {0}.", logo));
                    }
                    else
                    {
                        //XImage xImage = XImage.FromFile(logo);
                        //gfx.DrawImage(xImage, 20, 20, 104, 60);
                    }
                    */
                    //MessageBox.Show("cuantas van : " + cuantasvan);
                    if (cuantasvan == 0)
                    {

                        x = 0;
                        /*
                        gfx.DrawLine(pen, 1, 20, 1, 703);
                        gfx.DrawLine(pen, x + 50, 20, x + 50, 703);
                        gfx.DrawLine(pen, x + 100, 20, x + 100, 703);
                        gfx.DrawLine(pen, x + 150, 20, x + 150, 703);
                        gfx.DrawLine(pen, x + 200, 20, x + 200, 703);
                        gfx.DrawLine(pen, x + 250, 20, x + 250, 703);
                        gfx.DrawLine(pen, x + 300, 20, x + 300, 703);
                        gfx.DrawLine(pen, x + 350, 20, x + 350, 703);
                        gfx.DrawLine(pen, x + 400, 20, x + 400, 703);
                        gfx.DrawLine(pen, x + 450, 20, x + 450, 703);
                        gfx.DrawLine(pen, x + 500, 20, x + 500, 703);
                        gfx.DrawLine(pen, x + 550, 20, x + 550, 703);
                        gfx.DrawLine(pen, x + 600, 20, x + 600, 703);
                        gfx.DrawLine(pen, x + 650, 20, x + 650, 703);


                        gfx.DrawLine(pen, 45, 250, 45, 703);
                        gfx.DrawLine(pen, 87, 250, 87, 703);
                        gfx.DrawLine(pen, 150, 250, 150, 703);
                        gfx.DrawLine(pen, 291, 250, 291, 703);
                        gfx.DrawLine(pen, 381, 250, 381, 703);
                        gfx.DrawLine(pen, 461, 250, 461, 703);
                        gfx.DrawLine(pen, 571, 250, 571, 703);
                        */






                        //DATOS de encabezado
                        gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                        gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                        gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                        //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                        //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                        gfx.DrawString("FACTURA", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                        y = 85;

                        //DATOS de encabezado CLIENTE
                        gfx.DrawRectangle(pen, 13, 85, 330, 65);
                        gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                        //A LA DERECHA
                        gfx.DrawRectangle(pen, 400, 85, 180, 55);
                        gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("FACTURA Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                        //ENCABEZADOS PARTIDAS
                        y = y + 70;
                        gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                        gfx.DrawString("DESCRIPCIÓN DE LA MERCANCÍA", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("UNIDADES", ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PRECIO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("IMPORTE TOTAL", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                        gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                        gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                        gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                        //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                        y = y + 20;
                        gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                        gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                        gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                        gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                        gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                        gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                        gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                        gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                        gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                        gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                        //gfx.DrawLine(pen, 400, y - 2, 400, y + 12);


                        // sub encabezados pequeños

                        // fin de sub encabezados pequeños
                        y = y + 15;

                    }

                    y = y + 13;
                    x = 100;
                    inity = 70;
                    cuantasvan = 1;
                    //gfx.DrawString(cuantasvan + " partidad:" + nife, SIETE, XBrushes.Black, new XRect(200, 49, 100, y), XStringFormats.Center);
                    //gfx.DrawString("Pag: " + conto + " de " + contgral, SIETE, XBrushes.Black, new XRect(500, 25, 100, 25), XStringFormats.Center);

                    #endregion headerproform

             #region details


                    //PRODUCTOS

                    //Un elemento de la lista puede cambiar su valor de manera similar usando el índice combinado con el operador de asignación.
                    //Por ejemplo, para cambiar el color de verde a mamey:
                    //ListaColores[2] = "mamey";




                    //PRODUCTOS

                    int cuentaprod = rowids.Count();

                    /*
                    contadoarr = 0;
                    foreach (string dato in arrfilas)
                    {
                        MessageBox.Show(i + ") Ver contenido del array editable: " + arrfilasedit[contadoarr]);
                        arrfilasedit[contadoarr] = "" + dato.Replace("no","si");
                        contadoarr = contadoarr + 1;
                    }
                    contadoarr = 0;

                    foreach (string dato in arrfilas)
                    {
                        MessageBox.Show(i + ") Ver contenido del array editable: " + arrfilasedit[contadoarr]);
                        contadoarr = contadoarr + 1;
                    }
                    */
                    contadoarr = 0;

                    query = "SELECT * FROM rowsipl WHERE ord ='" + idinvo.Text + "' ORDER BY id ASC;";
                    cm = new SqlCommand(query, con);
                    da = new SqlDataAdapter(cm);
                    dt = new DataTable();
                    da.Fill(dt);
                    cuenta = dt.Rows.Count;
                    if (cuenta > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            foreach (string dato in arrfilas)
                            {

                                if (contadoarr <= cuentaprod - 1)
                                {
                                    //MessageBox.Show( contadoarr +"Buscando porque no recorre ?: " + row["id"].ToString() + " array: " + arrfilasedit[contadoarr]);
                                    //MessageBox.Show(i +") Antes de eval: " + arrfilasedit[contadoarr]);
                                    //string[] datos = arrfilasedit[contadoarr].Split(new char[] { '|' });
                                    //Si ocupa el segundo es el índice 1 pues el primero es índice 0.
                                    //string idero = datos[0];
                                    //string yesno = datos[1];
                                    //MessageBox.Show("Buscando: " + arrfilas[contadoarr] +  " contra: " + idero +" que debe ser igual a: " + row["id"].ToString());
                                    // MessageBox.Show("estado actual: " + dato);

                                    //MessageBox.Show("Comparando: " + row["id"].ToString() + "|no"  + " " + arrfilasedit[contadoarr] );
                                    if (row["id"].ToString() + "|no" == arrfilasedit[contadoarr])
                                    {
                                        arrfilasedit[contadoarr] = "" + dato.Replace("no", "si");
                                        //  arrfilasedit[contadoarr] = "" + row["id"].ToString() + "|si";
                                        //MessageBox.Show(" -se imprimio el " + row["id"].ToString() + " - debe decir yes: "  + arrfilasedit[contadoarr]);

                                        string querya = "SELECT * FROM artsipl WHERE clave ='" + row["clave"] + "';";
                                        SqlCommand cma = new SqlCommand(querya, con);
                                        SqlDataAdapter daa = new SqlDataAdapter(cma);
                                        DataTable dta = new DataTable();
                                        daa.Fill(dta);
                                        int cuentaa = dta.Rows.Count;
                                        if (cuentaa > 0)
                                        {
                                            foreach (DataRow rowa in dta.Rows)
                                            {
                                                size = "" + rowa["size"]; sizel = "" + rowa["sizel"];
                                                caja = "" + rowa["caja"]; mtscaja = "" + rowa["mtscaja"]; kgspiece = "" + rowa["kgspiece"];
                                                kgscaja = "" + rowa["kgscaja"]; sku = "" + rowa["size"];
                                                pallet = "" + rowa["pallet"];
                                                units = "" + rowa["ume"];
                                            }

                                        }
                                        cma.Dispose(); daa.Dispose(); dta.Dispose();



                                        ctns = 0;
                                        sqm = 0;
                                        tctns = 0;
                                        tsqm = 0;
                                        cajad = 0;
                                        mtsd = 0;
                                        pallets = 0;
                                        pieces = 0;

                                        try { ctns = double.Parse("" + caja); }
                                        catch { ctns = 0; }

                                        try { sqm = double.Parse("" + mtscaja); }
                                        catch { sqm = 0; }

                                        try { pallets = double.Parse("" + pallet); }
                                        catch { pallets = 0; }


                                        try
                                        {
                                            pieces = (ctns * double.Parse("" + row["cant"])) / sqm;
                                        }
                                        catch
                                        {
                                            pieces = 0;
                                        }


                                        try
                                        {
                                            tctns = double.Parse("" + row["cant"]) / sqm;
                                        }
                                        catch { tctns = 0; }

                                        try
                                        {
                                            tsqm = sqm * tctns;
                                        }
                                        catch { tsqm = 0; }
                                        if (row["pallets"].ToString() != "")
                                        {
                                            pallets = double.Parse("" + row["pallets"]);
                                        }
                                        else
                                        {
                                            try
                                            {
                                                pallets = double.Parse("" + row["cant"]) / pallets;
                                            }
                                            catch { pallets = 1; }
                                            //if (pallets < 1) { pallets = 1; }
                                            if (pallets.ToString() == "∞") { pallets = 0; }
                                        }
                                        try
                                        {
                                            cajad = double.Parse("" + row["cant"]) / sqm;
                                        }
                                        catch { cajad = 0; }



                                        if (contadoarr < cuentaprod - 1)
                                        {
                                            cajastot = cajastot + tctns;
                                            palletstot = palletstot + pallets;
                                            piezastot = piezastot + pieces;
                                        }
                                        //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                                        //gfx.DrawString("" + size +"X"+ sizel, ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString(""+ row["clave"].ToString(), ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("1", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        tf.Alignment = XParagraphAlignment.Right;
                                        gfx.DrawString("" + size + "X" + sizel, ocho, XBrushes.Black, new XRect(20, y, 100, 25), XStringFormats.TopLeft);
                                        gfx.DrawString("" + row["clave"].ToString(), SIETE, XBrushes.Black, new XRect(60, y, 100, 25), XStringFormats.TopLeft);
                                        gfx.DrawString("1", ocho, XBrushes.Black, new XRect(163, y, 100, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + pallets.ToString("n2"), ocho, XBrushes.Black, new XRect(175, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + Convert.ToInt32(tctns).ToString("N0"), ocho, XBrushes.Black, new XRect(235, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(295, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + Convert.ToInt32(pieces).ToString("N0"), ocho, XBrushes.Black, new XRect(340, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString(csymbol_l + "" + "" + row["pu"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                                        tf.DrawString(csymbol_l + "" + row["importe"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);

                                        y = y + 13;
                                        cuantasvan = cuantasvan + 1;
                                        cuantastotoal = cuantastotoal + 1;
                                        //SECCION PARA PODER AGREGAR PAGINAS NUEVAS AL PDF
                                        if (cuantasvan == 43)
                                        {
                                            if (cuantastotoal == totalpartidas)
                                            {

                                            }
                                            else
                                            {
                                                page = document.AddPage();
                                                page.Orientation = PdfSharp.PageOrientation.Portrait;
                                                page.Size = PdfSharp.PageSize.Letter;
                                                gfx = XGraphics.FromPdfPage(page);
                                                tf = new XTextFormatter(gfx);
                                                cuantasvan = 1;
                                                //DATOS de encabezado
                                                gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                                                //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                                                //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("FACTURA", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                                                y = 85;

                                                //DATOS de encabezado CLIENTE
                                                gfx.DrawRectangle(pen, 13, 85, 330, 65);
                                                gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                                                //A LA DERECHA
                                                gfx.DrawRectangle(pen, 400, 85, 180, 55);
                                                gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("FACTURA Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                                                //ENCABEZADOS PARTIDAS
                                                y = y + 70;
                                                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                                                gfx.DrawString("DESCRIPCIÓN DE LA MERCANCÍA", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("UNIDADES", ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PRECIO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("IMPORTE TOTAL", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                                                gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                                                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                                                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                                                //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                                                y = y + 20;
                                                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                                                gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                                                gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                                                gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                                                gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                                                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                                                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                                                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                                                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                                                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                                                y = y + 15;
                                                y = y + 13;
                                                x = 100;
                                                inity = 70;
                                                cuantasvan = 1;
                                            }
                                        }
                                        //FIN DE SECCION PARA PODER AGREGAR PAGINAS NUEVAS AL PDF
                                    }
                                    contadoarr = contadoarr + 1;
                                }
                            }     // for each que recorre el array maestro
                            contadoarr = 0;
                        }// for each que recorre el bloque de productos en la factura desde la base
                    }
                    cm.Dispose(); da.Dispose(); dt.Dispose();
                    contadoarr = 0;
                    #endregion details

             #region footer
                } // ESTE ES EL FOR EACH MAESTRO DE LOS REGISTROS DE PIEZAS Y SERVICIOS

                //HOJA ESPECIAL DE RESUMEN
                    page = document.AddPage();
                    page.Orientation = PdfSharp.PageOrientation.Portrait;
                    page.Size = PdfSharp.PageSize.Letter;
                    gfx = XGraphics.FromPdfPage(page);
                    tf = new XTextFormatter(gfx);

                    cuantasvan = 1;
                    //DATOS de encabezado
                    gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                    gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                    gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                    //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                    //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                    gfx.DrawString("FACTURA", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                    y = 85;

                    //DATOS de encabezado CLIENTE
                    gfx.DrawRectangle(pen, 13, 85, 330, 65);
                    gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                    //A LA DERECHA
                    gfx.DrawRectangle(pen, 400, 85, 180, 55);
                    gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("FACTURA Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                    //ENCABEZADOS PARTIDAS
                    y = y + 70;
                    gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                    gfx.DrawString("GASTOS INDIRECTOS", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("CANTIDAD", ocho, XBrushes.Black, new XRect(355, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PRECIO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("IMPORTE TOTAL", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                    gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                    gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                    gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                /*
                y = y + 20;
                    gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                    gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);
                

                    gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                    gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                    gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                    gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                    gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                    gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                    gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                    gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                    */

                y = y + 13;
                // SERVICIOS
                query = "SELECT *  FROM rowsservpl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                        if (cuenta > 0)
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                tf.Alignment = XParagraphAlignment.Right;
                                gfx.DrawString("" + row["descrip"], ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                                tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(298, y, 100, 25), XStringFormats.TopLeft);
                                tf.DrawString(csymbol_l + "" + "" + row["cu"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                                tf.DrawString(csymbol_l + "" + row["total"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);
                                y = y + 13;
                            }
                        }
                        cm.Dispose(); da.Dispose(); dt.Dispose();

                y = y + 15;
                y = y + 13;
                x = 100;
                inity = 70;

                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(358, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO NETO", ocho, XBrushes.Black, new XRect(415, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO BRUTO", ocho, XBrushes.Black, new XRect(510, y, 50, 25), XStringFormats.TopLeft);

                

                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);



                double metrostotals = 0;
                query = "SELECT SUM(convert(numeric(18, 6), replace(cant, ',', ''))) as totals FROM rowsipl  where ord = '" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        metrostotals = double.Parse("" + row["totals"]);
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                cajastot = cajastot + tctns; palletstot = palletstot + pallets; piezastot = piezastot + pieces;

                y = y + 20;
                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("" + palletstot.ToString("N2"), ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + cajastot.ToString("N0"), ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + metrostotals.ToString("N2"), ocho, XBrushes.Black, new XRect(305, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + piezastot.ToString("N0"), ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString(""+ netofact, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + brutosfact, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);



                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                y = y + 20;
                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("CONTENEDOR(ES)", ocho, XBrushes.Black, new XRect(197, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PRECINTO", ocho, XBrushes.Black, new XRect(290, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO NETO EN KGS", ocho, XBrushes.Black, new XRect(400, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO BRUTO EN KGS", ocho, XBrushes.Black, new XRect(500, y, 50, 25), XStringFormats.TopLeft);

                gfx.DrawLine(pen, 288, y - 2, 288, y + 12);
                gfx.DrawLine(pen, 396, y - 2, 396, y + 12);
                gfx.DrawLine(pen, 498, y - 2, 498, y + 12);



                y = y + 17;
                string queryc = "SELECT * FROM containersipl WHERE ord ='" + id + "' ORDER BY id ASC;";
                SqlCommand cmc = new SqlCommand(queryc, con);
                SqlDataAdapter dac = new SqlDataAdapter(cmc);
                DataTable dtc = new DataTable();
                dac.Fill(dtc);
                int cuentac = dtc.Rows.Count;
                if (cuentac > 0)
                {
                    foreach (DataRow rowc in dtc.Rows)
                    {
                        gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                        gfx.DrawString("" + rowc["container"], ocho, XBrushes.Black, new XRect(197, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + rowc["precinto"], ocho, XBrushes.Black, new XRect(293, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + double.Parse("" + rowc["pesoneto"]), ocho, XBrushes.Black, new XRect(403, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + double.Parse("" + rowc["pesobruto"]), ocho, XBrushes.Black, new XRect(503, y, 50, 25), XStringFormats.TopLeft);

                        gfx.DrawLine(pen, 288, y - 2, 288, y + 12);
                        gfx.DrawLine(pen, 396, y - 2, 396, y + 12);
                        gfx.DrawLine(pen, 498, y - 2, 498, y + 12);

                        y = y + 17;
                    }
                }
                cmc.Dispose(); dac.Dispose(); dtc.Dispose();



                y = y + 20;
                gfx.DrawRectangle(pen, 30, y - 2,565, 14);
                gfx.DrawString("DESTINO: ", ocho, XBrushes.Black, new XRect(33, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + origdest, ocho, XBrushes.Black, new XRect(85, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawLine(pen, 80, y - 2, 80, y + 12);

                y = y + 25;
                tf.Alignment = XParagraphAlignment.Right;
                gfx.DrawRectangle(pen, 376, y - 2, 220, 14);
                gfx.DrawString("TOTAL FACTURA", dieznegra, XBrushes.Black, new XRect(380, y, 50, 25), XStringFormats.TopLeft);
                tf.DrawString("" + csymbol_l + "" + tot + "" + csymbol_r, dieznegra, XBrushes.Black, new XRect(500, y, 90, 25), XStringFormats.TopLeft);
                gfx.DrawLine(pen, 478, y - 2, 478, y + 12);

                y = y + 25;
                tf.Alignment = XParagraphAlignment.Justify;
                gfx.DrawRectangle(pen, 30, y - 2, 565, 120);
                gfx.DrawString("Observaciones: ", ocho, XBrushes.Black, new XRect(33, y, 50, 25), XStringFormats.TopLeft);
                y = y + 13;
                tf.DrawString("" + obs1 + "\r\n" + obs2 + "\r\n" + obs3 + "\r\n" + obs4 + "\r\n" + obs5, CINCO, XBrushes.Black, new XRect(40, y, 548, 100), XStringFormats.TopLeft);



                #endregion footer



                #region initpl            
                //INICIA PACKING LIST
                page = document.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Portrait;
                page.Size = PdfSharp.PageSize.Letter;
                gfx = XGraphics.FromPdfPage(page);
                tf = new XTextFormatter(gfx);


                id = ""; numees = ""; folio = ""; empresa = ""; idcli = ""; nomcli = ""; callecli = "";
                numcli = ""; numclii = ""; colcli = ""; muncli = ""; edocli = ""; paiscli = ""; fecha = "";
                albaran = ""; origdest = ""; tot = ""; currency = ""; cpcli = ""; nifcli = "";
                ide = ""; nome = ""; nife = ""; callee = ""; nume = ""; numie = ""; cole = ""; cde = "";
                estadoe = ""; paise = ""; cpe = "";
                netofact = ""; brutosfact = "";
                obs1 = ""; obs2 = ""; obs3 = ""; obs4 = ""; obs5 = "";
                csymbol_l = "";
                csymbol_r = "";

                size = ""; sizel = "";
                caja = ""; mtscaja = ""; kgspiece = "";
                kgscaja = ""; sku = "";
                pallet = "";
                units = "";


                kgscontainer = 0; kgsacums = 0;
                cuantasvan = 0; cuantastotoal = 0;

                mtspartidas = 0; mtsacum = 0; kgs1caja = 0; mts1caja = 0; tonelaje = 0; tonelajecontainer = 0;
                tarimaskgs = 0; kgstar = 0; kilosXpallets = 0;
                netofull = 0; brutofull = 0;
                cajastot = 0; palletstot = 0; piezastot = 0;
                cellrow1 = ""; cellrow2 = "";
                inity = 70; y = 25; x = 100;

                sqm = 0;
                tctns = 0;
                tsqm = 0;
                cajad = 0;
                mtsd = 0;
                pallets = 0;
                pieces = 0;
                ctns = 0;
                totalpartidas = 0;

                cuantasvan = 0;
                contadoarr = 0;

                #endregion initpl

                #region headerproformpl

                //concentrado.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //concentrado.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);



                query = "SELECT * FROM invoicespl WHERE id =" + idinvo.Text + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {

                        id = "" + row["id"]; folio = "" + row["folio"]; empresa = "" + row["empresa"]; idcli = "" + row["idcli"];
                        nomcli = "" + row["nomcli"];
                        callecli = "" + row["callecli"];
                        numcli = "" + row["numcli"]; numclii = "" + row["numclii"]; colcli = "" + row["colcli"];
                        muncli = "" + row["muncli"];
                        edocli = "" + row["edocli"]; paiscli = "" + row["paiscli"];
                        fecha = "" + row["fecha"];
                        origdest = "" + row["origdest"]; tot = "" + row["tot"];
                        currency = "" + row["currency"];

                        obs1 = "" + row["obs1"];
                        obs2 = "" + row["obs2"];
                        obs3 = "" + row["obs3"];
                        obs4 = "" + row["obs4"];
                        obs5 = "" + row["obs5"];

                        albaran = "" + row["albaran"];
                        numees = "" + row["number"];
                        netofact = "" + row["pesoneto"]; brutosfact = "" + row["pesobruto"];
                    }
                }
                else
                {
                    MessageBox.Show("Orden no existe");
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT * FROM empresasipl WHERE id =" + empresa + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT id,cp,nif FROM clientesipl WHERE id =" + idcli + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        cpcli = "" + row["cp"];
                        nifcli = "" + row["nif"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                if (currency == "USD") { csymbol_l = "US $"; }
                if (currency == "EUR") { csymbol_r = " €"; }




                query = "SELECT id FROM rowsipl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    totalpartidas = totalpartidas + cuenta;
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                query = "SELECT id  FROM rowsservpl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    totalpartidas = totalpartidas + cuenta;
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                //MessageBox.Show(" total de partidas para pintar: " + totalpartidas);
                cuenta = 0;

                //creacion de las listas para productos y servicios y saber cuales estan inmpresos y cuales no

                conteo = 0;
                contadoarr = 0;
                rowids.Clear();
                List<string> rowidspl = new List<string>();
                query = "SELECT id FROM rowsipl WHERE ord ='" + idinvo.Text + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        rowidspl.Add(row["id"] + "|no");
                        conteo = conteo + 1;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                /*
                query = "SELECT id FROM rowsservpl WHERE ord ='" + idinvo.Text + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        rowidserv.Add(row["id"] + "|no");
                        conteo = conteo + 1;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                */

                //arrfilas[conteo] = "no|" + row["id"];
                //arrfilas.
                //arrfilas = rowidspl.ToArray();
                String[] arrfilaspl = rowidspl.ToArray();
                String[] arrfilaseditpl = rowidspl.ToArray();
                //arrfilasedit = rowids.ToArray();
                //FIN DE LA CREACION creacion de las listas para productos y servicios y saber cuales estan inmpresos y cuales no


                //FOR EACH MAESTRO DE PIEZAS Y SERVICIOS

                cajastot = 0;
                palletstot = 0;
                piezastot = 0;
                for (int i = 1; i <= totalpartidas; i++)
                {
                    /*
                    string logo = @"" + Path.GetDirectoryName(Application.ExecutablePath) + @"\" + "logo.jpg";
                    if (!File.Exists(@"" + logo))
                    {
                        //throw new FileNotFoundException(String.Format("No se encuentra el Logo {0}.", logo));
                    }
                    else
                    {
                        //XImage xImage = XImage.FromFile(logo);
                        //gfx.DrawImage(xImage, 20, 20, 104, 60);
                    }
                    */
                    //MessageBox.Show("cuantas van : " + cuantasvan);
                    if (cuantasvan == 0)
                    {

                        x = 0;
                        /*
                        gfx.DrawLine(pen, 1, 20, 1, 703);
                        gfx.DrawLine(pen, x + 50, 20, x + 50, 703);
                        gfx.DrawLine(pen, x + 100, 20, x + 100, 703);
                        gfx.DrawLine(pen, x + 150, 20, x + 150, 703);
                        gfx.DrawLine(pen, x + 200, 20, x + 200, 703);
                        gfx.DrawLine(pen, x + 250, 20, x + 250, 703);
                        gfx.DrawLine(pen, x + 300, 20, x + 300, 703);
                        gfx.DrawLine(pen, x + 350, 20, x + 350, 703);
                        gfx.DrawLine(pen, x + 400, 20, x + 400, 703);
                        gfx.DrawLine(pen, x + 450, 20, x + 450, 703);
                        gfx.DrawLine(pen, x + 500, 20, x + 500, 703);
                        gfx.DrawLine(pen, x + 550, 20, x + 550, 703);
                        gfx.DrawLine(pen, x + 600, 20, x + 600, 703);
                        gfx.DrawLine(pen, x + 650, 20, x + 650, 703);


                        gfx.DrawLine(pen, 45, 250, 45, 703);
                        gfx.DrawLine(pen, 87, 250, 87, 703);
                        gfx.DrawLine(pen, 150, 250, 150, 703);
                        gfx.DrawLine(pen, 291, 250, 291, 703);
                        gfx.DrawLine(pen, 381, 250, 381, 703);
                        gfx.DrawLine(pen, 461, 250, 461, 703);
                        gfx.DrawLine(pen, 571, 250, 571, 703);
                        */






                        //DATOS de encabezado
                        gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                        gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                        gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                        //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                        //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                        gfx.DrawString("PACKING LIST", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                        y = 85;

                        //DATOS de encabezado CLIENTE
                        gfx.DrawRectangle(pen, 13, 85, 330, 65);
                        gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                        //A LA DERECHA
                        gfx.DrawRectangle(pen, 400, 85, 180, 55);
                        gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PACKING LIST Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                        //ENCABEZADOS PARTIDAS
                        y = y + 70;
                        gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                        gfx.DrawString("DESCRIPCIÓN DE LA MERCANCÍA", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("UNIDADES", ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO BRUTO", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                        gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                        gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                        gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                        //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                        y = y + 20;
                        gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                        gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                        gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                        gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                        gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                        gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                        gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                        gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                        gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                        gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                        //gfx.DrawLine(pen, 400, y - 2, 400, y + 12);


                        // sub encabezados pequeños

                        // fin de sub encabezados pequeños
                        y = y + 15;

                    }

                    y = y + 13;
                    x = 100;
                    inity = 70;
                    cuantasvan = 1;
                    //gfx.DrawString(cuantasvan + " partidad:" + nife, SIETE, XBrushes.Black, new XRect(200, 49, 100, y), XStringFormats.Center);
                    //gfx.DrawString("Pag: " + conto + " de " + contgral, SIETE, XBrushes.Black, new XRect(500, 25, 100, 25), XStringFormats.Center);

                    #endregion headerproformpl

                    #region detailspl


                    //PRODUCTOS

                    //Un elemento de la lista puede cambiar su valor de manera similar usando el índice combinado con el operador de asignación.
                    //Por ejemplo, para cambiar el color de verde a mamey:
                    //ListaColores[2] = "mamey";




                    //PRODUCTOS

                    int cuentaprod = rowidspl.Count();

                    /*
                    contadoarr = 0;
                    foreach (string dato in arrfilas)
                    {
                        MessageBox.Show(i + ") Ver contenido del array editable: " + arrfilasedit[contadoarr]);
                        arrfilasedit[contadoarr] = "" + dato.Replace("no","si");
                        contadoarr = contadoarr + 1;
                    }
                    contadoarr = 0;

                    foreach (string dato in arrfilas)
                    {
                        MessageBox.Show(i + ") Ver contenido del array editable: " + arrfilasedit[contadoarr]);
                        contadoarr = contadoarr + 1;
                    }
                    */
                    //contadoarr = 0;

                    query = "SELECT * FROM rowsipl WHERE ord ='" + idinvo.Text + "' ORDER BY container,id ASC;";
                    cm = new SqlCommand(query, con);
                    da = new SqlDataAdapter(cm);
                    dt = new DataTable();
                    da.Fill(dt);
                    cuenta = dt.Rows.Count;
                    if (cuenta > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            foreach (string dato in arrfilaspl)
                            {

                                if (contadoarr <= cuentaprod - 1)
                                {
                                    //MessageBox.Show( contadoarr +"Buscando porque no recorre ?: " + row["id"].ToString() + " array: " + arrfilasedit[contadoarr]);
                                    //MessageBox.Show(i +") Antes de eval: " + arrfilasedit[contadoarr]);
                                    //string[] datos = arrfilasedit[contadoarr].Split(new char[] { '|' });
                                    //Si ocupa el segundo es el índice 1 pues el primero es índice 0.
                                    //string idero = datos[0];
                                    //string yesno = datos[1];
                                    //MessageBox.Show("Buscando: " + arrfilas[contadoarr] +  " contra: " + idero +" que debe ser igual a: " + row["id"].ToString());
                                    // MessageBox.Show("estado actual: " + dato);

                                    //MessageBox.Show("Comparando: " + row["id"].ToString() + "|no"  + " " + arrfilaseditpl[contadoarr] );
                                    if (row["id"].ToString() + "|no" == arrfilaseditpl[contadoarr])
                                    {
                                        arrfilaseditpl[contadoarr] = "" + dato.Replace("no", "si");
                                        //  arrfilasedit[contadoarr] = "" + row["id"].ToString() + "|si";
                                        //MessageBox.Show(" -se imprimio el " + row["id"].ToString() + " - debe decir yes: "  + arrfilasedit[contadoarr]);

                                        string querya = "SELECT * FROM artsipl WHERE clave ='" + row["clave"] + "';";
                                        SqlCommand cma = new SqlCommand(querya, con);
                                        SqlDataAdapter daa = new SqlDataAdapter(cma);
                                        DataTable dta = new DataTable();
                                        daa.Fill(dta);
                                        int cuentaa = dta.Rows.Count;
                                        if (cuentaa > 0)
                                        {
                                            foreach (DataRow rowa in dta.Rows)
                                            {
                                                size = "" + rowa["size"]; sizel = "" + rowa["sizel"];
                                                caja = "" + rowa["caja"]; mtscaja = "" + rowa["mtscaja"]; kgspiece = "" + rowa["kgspiece"];
                                                kgscaja = "" + rowa["kgscaja"]; sku = "" + rowa["size"];
                                                pallet = "" + rowa["pallet"];
                                                units = "" + rowa["ume"];
                                            }

                                        }
                                        cma.Dispose(); daa.Dispose(); dta.Dispose();



                                        ctns = 0;
                                        sqm = 0;
                                        tctns = 0;
                                        tsqm = 0;
                                        cajad = 0;
                                        mtsd = 0;
                                        pallets = 0;
                                        pieces = 0;

                                        try { ctns = double.Parse("" + caja); }
                                        catch { ctns = 0; }

                                        try { sqm = double.Parse("" + mtscaja); }
                                        catch { sqm = 0; }

                                        try { pallets = double.Parse("" + pallet); }
                                        catch { pallets = 0; }


                                        try
                                        {
                                            pieces = (ctns * double.Parse("" + row["cant"])) / sqm;
                                        }
                                        catch
                                        {
                                            pieces = 0;
                                        }


                                        try
                                        {
                                            tctns = double.Parse("" + row["cant"]) / sqm;
                                        }
                                        catch { tctns = 0; }

                                        try
                                        {
                                            tsqm = sqm * tctns;
                                        }
                                        catch { tsqm = 0; }
                                        if (row["pallets"].ToString() != "")
                                        {
                                            pallets = double.Parse("" + row["pallets"]);
                                        }
                                        else
                                        {
                                            try
                                            {
                                                pallets = double.Parse("" + row["cant"]) / pallets;
                                            }
                                            catch { pallets = 1; }
                                            //if (pallets < 1) { pallets = 1; }
                                            if (pallets.ToString() == "∞") { pallets = 0; }
                                        }
                                        try
                                        {
                                            cajad = double.Parse("" + row["cant"]) / sqm;
                                        }
                                        catch { cajad = 0; }



                                        if (contadoarr < cuentaprod - 1)
                                        {
                                            cajastot = cajastot + tctns;
                                            palletstot = palletstot + pallets;
                                            piezastot = piezastot + pieces;
                                        }
                                        //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                                        //gfx.DrawString("" + size +"X"+ sizel, ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString(""+ row["clave"].ToString(), ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("1", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        tf.Alignment = XParagraphAlignment.Right;
                                        gfx.DrawString("" + size + "X" + sizel, ocho, XBrushes.Black, new XRect(20, y, 100, 25), XStringFormats.TopLeft);
                                        gfx.DrawString("" + row["clave"].ToString(), SIETE, XBrushes.Black, new XRect(60, y, 100, 25), XStringFormats.TopLeft);
                                        gfx.DrawString("1", ocho, XBrushes.Black, new XRect(163, y, 100, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + pallets.ToString("n2"), ocho, XBrushes.Black, new XRect(175, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + Convert.ToInt32(tctns).ToString("N0"), ocho, XBrushes.Black, new XRect(235, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(295, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + Convert.ToInt32(pieces).ToString("N0"), ocho, XBrushes.Black, new XRect(340, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + row["pesoneto"], ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + row["pesobruto"], ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);

                                        y = y + 13;



                                        // INIT bloque para poder imprimir el totalizador de container
                                        string bsql = "SELECT TOP (1) * FROM rowsipl WHERE ord = '" + row["ord"].ToString() + "' AND container= '" + row["container"].ToString() + "' ORDER BY id desc";
                                        SqlCommand cmab = new SqlCommand(bsql, con);
                                        SqlDataAdapter daab = new SqlDataAdapter(cmab);
                                        DataTable dtab = new DataTable();
                                        daab.Fill(dtab);
                                        int cuentaab = dtab.Rows.Count;

                                        if (cuentaab > 0)
                                        {
                                            foreach (DataRow rowaa in dtab.Rows)
                                            {
                                                if (rowaa["id"].ToString() + "" == "" + row["id"].ToString())
                                                {
                                                    string bsqlc = "SELECT precinto,pesoneto,pesobruto,container FROM containersipl WHERE ord = '" + row["ord"].ToString() + "' AND container= '" + row["container"].ToString() + "' ORDER BY id desc";
                                                    SqlCommand cmabc = new SqlCommand(bsqlc, con);
                                                    SqlDataAdapter daabc = new SqlDataAdapter(cmabc);
                                                    DataTable dtabc = new DataTable();
                                                    daabc.Fill(dtabc);
                                                    int cuentaabc = dtabc.Rows.Count;

                                                    if (cuentaabc > 0)
                                                    {
                                                        foreach (DataRow rowaac in dtabc.Rows)
                                                        {
                                                            //CONTENEDOR: RNHC6213642 PRECINTO: SIN PRECINTO              48,435.36   49,259.76

                                                            //MessageBox.Show("Imprimir remate totalizador del container: " + rowaa["id"].ToString() + " " + row["container"].ToString());


                                                            gfx.DrawRectangle(peng, 13, y - 2, 580, 14);
                                                            gfx.DrawLine(peng, 190, y - 2, 190, y + 12);
                                                            gfx.DrawLine(peng, 400, y - 2, 400, y + 12);
                                                            gfx.DrawLine(peng, 490, y - 2, 490, y + 12);


                                                            tf.Alignment = XParagraphAlignment.Right;
                                                            gfx.DrawString("CONTAINER: " + rowaac["container"].ToString(), ocho, XBrushes.Black, new XRect(20, y, 100, 25), XStringFormats.TopLeft);
                                                            //gfx.DrawString("", SIETE, XBrushes.Black, new XRect(60, y, 100, 25), XStringFormats.TopLeft);
                                                            gfx.DrawString("", ocho, XBrushes.Black, new XRect(163, y, 100, 25), XStringFormats.TopLeft);
                                                            tf.DrawString("PRECINTO: " + rowaac["precinto"].ToString(), ocho, XBrushes.Black, new XRect(100, y, 200, 25), XStringFormats.TopLeft);
                                                            //tf.DrawString("" + Convert.ToInt32(tctns).ToString("N0"), ocho, XBrushes.Black, new XRect(235, y, 50, 25), XStringFormats.TopLeft);
                                                            //tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(295, y, 50, 25), XStringFormats.TopLeft);
                                                            //tf.DrawString("" + Convert.ToInt32(pieces).ToString("N0"), ocho, XBrushes.Black, new XRect(340, y, 50, 25), XStringFormats.TopLeft);
                                                            tf.DrawString("" + rowaac["pesoneto"].ToString(), ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                                                            tf.DrawString("" + rowaac["pesobruto"].ToString(), ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);
                                                        }

                                                    }

                                                    y = y + 22;
                                                    cuantasvan = cuantasvan + 1;
                                                }
                                            }
                                        }
                                        // FIN bloque para poder imprimir el totalizador de container


                                        cuantasvan = cuantasvan + 1;
                                        cuantastotoal = cuantastotoal + 1;
                                        //SECCION PARA PODER AGREGAR PAGINAS NUEVAS AL PDF
                                        if (cuantasvan == 42)
                                        {
                                            if (cuantastotoal == totalpartidas)
                                            {

                                            }
                                            else
                                            {
                                                page = document.AddPage();
                                                page.Orientation = PdfSharp.PageOrientation.Portrait;
                                                page.Size = PdfSharp.PageSize.Letter;
                                                gfx = XGraphics.FromPdfPage(page);
                                                tf = new XTextFormatter(gfx);
                                                cuantasvan = 1;
                                                //DATOS de encabezado
                                                gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                                                //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                                                //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("PACKING LIST", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                                                y = 85;

                                                //DATOS de encabezado CLIENTE
                                                gfx.DrawRectangle(pen, 13, 85, 330, 65);
                                                gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                                                //A LA DERECHA
                                                gfx.DrawRectangle(pen, 400, 85, 180, 55);
                                                gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("FACTURA Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                                                //ENCABEZADOS PARTIDAS
                                                y = y + 70;
                                                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                                                gfx.DrawString("DESCRIPCIÓN DE LA MERCANCÍA", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("UNIDADES", ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO BRUTO", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                                                gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                                                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                                                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                                                //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                                                y = y + 20;
                                                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                                                gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                                                gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                                                gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                                                gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                                                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                                                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                                                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                                                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                                                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                                                y = y + 15;
                                                y = y + 13;
                                                x = 100;
                                                inity = 70;
                                                cuantasvan = 1;
                                            }
                                        }
                                        //FIN DE SECCION PARA PODER AGREGAR PAGINAS NUEVAS AL PDF
                                    }
                                    contadoarr = contadoarr + 1;
                                }
                            }     // for each que recorre el array maestro
                            contadoarr = 0;
                        }// for each que recorre el bloque de productos en la factura desde la base
                    }
                    cm.Dispose(); da.Dispose(); dt.Dispose();
                    contadoarr = 0;
                    #endregion detailspl

                    #region footerpl
                } // ESTE ES EL FOR EACH MAESTRO DE LOS REGISTROS DE PIEZAS Y SERVICIOS

                //HOJA ESPECIAL DE RESUMEN
                page = document.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Portrait;
                page.Size = PdfSharp.PageSize.Letter;
                gfx = XGraphics.FromPdfPage(page);
                tf = new XTextFormatter(gfx);

                cuantasvan = 1;
                //DATOS de encabezado
                gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                gfx.DrawString("PACKING LIST - RESUMEN", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                y = 85;

                //DATOS de encabezado CLIENTE
                gfx.DrawRectangle(pen, 13, 85, 330, 65);
                gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                //A LA DERECHA
                gfx.DrawRectangle(pen, 400, 85, 180, 55);
                gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("FACTURA Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                //ENCABEZADOS PARTIDAS
                y = y + 70;
                /*
                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                gfx.DrawString("GASTOS INDIRECTOS", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("CANTIDAD", ocho, XBrushes.Black, new XRect(355, y, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("PRECIO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("IMPORTE TOTAL", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);

                */

                //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                /*
                y = y + 20;
                    gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                    gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                    gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                    gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                    gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                    gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                    gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                    gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                    gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                    gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                    */

                y = y + 13;
                // SERVICIOS
                /*
                query = "SELECT *  FROM rowsservpl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        tf.Alignment = XParagraphAlignment.Right;
                        gfx.DrawString("" + row["descrip"], ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                        tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(298, y, 100, 25), XStringFormats.TopLeft);
                        tf.DrawString(csymbol_l + "" + "" + row["cu"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                        tf.DrawString(csymbol_l + "" + row["total"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);
                        y = y + 13;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                */
                y = y + 15;
                y = y + 13;
                x = 100;
                inity = 70;

                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(358, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO NETO", ocho, XBrushes.Black, new XRect(415, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO BRUTO", ocho, XBrushes.Black, new XRect(510, y, 50, 25), XStringFormats.TopLeft);



                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);



                metrostotals = 0;
                query = "SELECT SUM(convert(numeric(18, 6), replace(cant, ',', ''))) as totals FROM rowsipl  where ord = '" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        metrostotals = double.Parse("" + row["totals"]);
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                cajastot = cajastot + tctns; palletstot = palletstot + pallets; piezastot = piezastot + pieces;

                y = y + 20;
                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("" + palletstot.ToString("N2"), ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + cajastot.ToString("N0"), ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + metrostotals.ToString("N2"), ocho, XBrushes.Black, new XRect(305, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + piezastot.ToString("N0"), ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + netofact, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + brutosfact, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);



                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                y = y + 60;
                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("CONTENEDOR(ES)", ocho, XBrushes.Black, new XRect(197, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PRECINTO", ocho, XBrushes.Black, new XRect(290, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO NETO EN KGS", ocho, XBrushes.Black, new XRect(400, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO BRUTO EN KGS", ocho, XBrushes.Black, new XRect(500, y, 50, 25), XStringFormats.TopLeft);

                gfx.DrawLine(pen, 288, y - 2, 288, y + 12);
                gfx.DrawLine(pen, 396, y - 2, 396, y + 12);
                gfx.DrawLine(pen, 498, y - 2, 498, y + 12);



                y = y + 17;
                queryc = "SELECT * FROM containersipl WHERE ord ='" + id + "' ORDER BY id ASC;";
                cmc = new SqlCommand(queryc, con);
                dac = new SqlDataAdapter(cmc);
                dtc = new DataTable();
                dac.Fill(dtc);
                cuentac = dtc.Rows.Count;
                if (cuentac > 0)
                {
                    foreach (DataRow rowc in dtc.Rows)
                    {
                        gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                        gfx.DrawString("" + rowc["container"], ocho, XBrushes.Black, new XRect(197, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + rowc["precinto"], ocho, XBrushes.Black, new XRect(293, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + double.Parse("" + rowc["pesoneto"]), ocho, XBrushes.Black, new XRect(403, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + double.Parse("" + rowc["pesobruto"]), ocho, XBrushes.Black, new XRect(503, y, 50, 25), XStringFormats.TopLeft);

                        gfx.DrawLine(pen, 288, y - 2, 288, y + 12);
                        gfx.DrawLine(pen, 396, y - 2, 396, y + 12);
                        gfx.DrawLine(pen, 498, y - 2, 498, y + 12);

                        y = y + 17;
                    }
                }
                cmc.Dispose(); dac.Dispose(); dtc.Dispose();



                y = y + 20;
                gfx.DrawRectangle(pen, 30, y - 2, 565, 14);
                gfx.DrawString("DESTINO: ", ocho, XBrushes.Black, new XRect(33, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + origdest, ocho, XBrushes.Black, new XRect(85, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawLine(pen, 80, y - 2, 80, y + 12);

                y = y + 25;
                /*
                tf.Alignment = XParagraphAlignment.Right;
                gfx.DrawRectangle(pen, 376, y - 2, 220, 14);
                gfx.DrawString("TOTAL FACTURA", dieznegra, XBrushes.Black, new XRect(380, y, 50, 25), XStringFormats.TopLeft);
                tf.DrawString("" + csymbol_l + "" + tot + "" + csymbol_r, dieznegra, XBrushes.Black, new XRect(500, y, 90, 25), XStringFormats.TopLeft);
                gfx.DrawLine(pen, 478, y - 2, 478, y + 12);
                */

                y = y + 25;

                tf.Alignment = XParagraphAlignment.Justify;
                gfx.DrawRectangle(pen, 30, y - 2, 565, 120);
                gfx.DrawString("Observaciones: ", ocho, XBrushes.Black, new XRect(33, y, 50, 25), XStringFormats.TopLeft);
                y = y + 13;
                tf.DrawString("" + obs1 + "\r\n" + obs2 + "\r\n" + obs3 + "\r\n" + obs4 + "\r\n" + obs5, CINCO, XBrushes.Black, new XRect(40, y, 548, 100), XStringFormats.TopLeft);



                #endregion footerpl


                #region opener

                try
                {
                    document.Save(@"" + paths);
                    // ...and start a viewer.
                    Process.Start(paths);
                }
                catch (Exception ex) { MessageBox.Show("Actualizando, presione el boton de nuevo\r\n" + ex.ToString()); }
                #endregion opener
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (idinvo.Text == "") { MessageBox.Show("No selecciono ningun documento para imprimir"); }
            else
            {
                #region initpl            
                DateTime hoy = DateTime.Now;
                logpesos.Text = "";
                string dates = hoy.Day.ToString() + "-" + hoy.Month.ToString() + "-" + hoy.Year.ToString() + "-" + hoy.Hour.ToString() + "-" + hoy.Minute.ToString() + "-" + hoy.Second.ToString();
                string paths = @"" + config.tempofiles + @"\INVOICE_PACKINGLIST_" + dates + ".pdf";
                // Create a new PDF document
                PdfDocument document = new PdfDocument();


                document.Info.Title = "Factura número: " + numee.Text;
                // Create an empty page Y PONERLE SUS PROPIEDADES
                PdfPage page = document.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Portrait;
                page.Size = PdfSharp.PageSize.Letter;
                // page = document.AddPage();
                // Get an XGraphics object for drawing
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XTextFormatter tf = new XTextFormatter(gfx);
                // FUENTES PARA EL DOCTO
                XFont dieznegra = new XFont("Arial", 10, XFontStyle.Bold);
                XFont catonegra = new XFont("Arial", 14, XFontStyle.Bold);

                XFont ocho = new XFont("Arial", 8, XFontStyle.Regular);
                XFont DOCE = new XFont("Arial", 12, XFontStyle.Regular);
                XFont ochoneg = new XFont("Arial", 8, XFontStyle.Bold);
                XFont SIETE = new XFont("Arial", 7, XFontStyle.Regular);
                XFont SEIS = new XFont("Arial", 6, XFontStyle.Regular);
                XFont CINCO = new XFont("Arial", 5, XFontStyle.Regular);
                XPen pen = new XPen(XColors.Black, 1);
                XPen peng = new XPen(XColors.Gray, 0.3);

                string id = "", numees = "", folio = "", empresa = "", idcli = "", nomcli = "", callecli = "", numcli = "", numclii = "", colcli = "", muncli = "", edocli = "", paiscli = "", fecha = "", albaran = "", origdest = "", tot = "", currency = "", cpcli = "", nifcli = "";
                string ide = "", nome = "", nife = "", callee = "", nume = "", numie = "", cole = "", cde = "", estadoe = "", paise = "", cpe = "";
                string netofact = "", brutosfact = "";
                string obs1 = "", obs2 = "", obs3 = "", obs4 = "", obs5 = "";
                string csymbol_l = "";
                string csymbol_r = "";

                string size = "", sizel = "";
                string caja = "", mtscaja = "", kgspiece = "";
                string kgscaja = "", sku = "";
                string pallet = "";
                string units = "";


                double kgscontainer = 0, kgsacums = 0;
                int cuantasvan = 0, cuantastotoal = 0;

                double mtspartidas = 0, mtsacum = 0, kgs1caja = 0, mts1caja = 0, tonelaje = 0, tonelajecontainer = 0;
                double tarimaskgs = 0, kgstar = 0, kilosXpallets = 0;
                double netofull = 0, brutofull = 0;
                double cajastot = 0, palletstot = 0, piezastot = 0;
                string cellrow1 = "", cellrow2 = "";
                int inity = 70, y = 25, x = 100;

                double sqm = 0;
                double tctns = 0;
                double tsqm = 0;
                double cajad = 0;
                double mtsd = 0;
                double pallets = 0;
                double pieces = 0;
                double ctns = 0;
                int totalpartidas = 0;

                #endregion initpl

                #region headerproformpl

                //concentrado.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //concentrado.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


                SqlConnection con = new SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT * FROM invoicespl WHERE id =" + idinvo.Text + ";";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {

                        id = "" + row["id"]; folio = "" + row["folio"]; empresa = "" + row["empresa"]; idcli = "" + row["idcli"];
                        nomcli = "" + row["nomcli"];
                        callecli = "" + row["callecli"];
                        numcli = "" + row["numcli"]; numclii = "" + row["numclii"]; colcli = "" + row["colcli"];
                        muncli = "" + row["muncli"];
                        edocli = "" + row["edocli"]; paiscli = "" + row["paiscli"];
                        fecha = "" + row["fecha"];
                        origdest = "" + row["origdest"]; tot = "" + row["tot"];
                        currency = "" + row["currency"];

                        obs1 = "" + row["obs1"];
                        obs2 = "" + row["obs2"];
                        obs3 = "" + row["obs3"];
                        obs4 = "" + row["obs4"];
                        obs5 = "" + row["obs5"];

                        albaran = "" + row["albaran"];
                        numees = "" + row["number"];
                        netofact = "" + row["pesoneto"]; brutosfact = "" + row["pesobruto"];
                    }
                }
                else
                {
                    MessageBox.Show("Orden no existe");
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT * FROM empresasipl WHERE id =" + empresa + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ide = "" + row["id"]; nome = "" + row["nom"]; nife = "" + row["nif"]; callee = "" + row["calle"]; nume = "" + row["num"]; numie = "" + row["numi"]; cole = "" + row["col"]; cde = "" + row["cd"]; estadoe = "" + row["estado"]; paise = "" + row["pais"]; cpe = "" + row["cp"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT id,cp,nif FROM clientesipl WHERE id =" + idcli + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        cpcli = "" + row["cp"];
                        nifcli = "" + row["nif"];
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                if (currency == "USD") { csymbol_l = "US $"; }
                if (currency == "EUR") { csymbol_r = " €"; }




                query = "SELECT id FROM rowsipl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    totalpartidas = totalpartidas + cuenta;
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                query = "SELECT id  FROM rowsservpl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    totalpartidas = totalpartidas + cuenta;
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                //MessageBox.Show(" total de partidas para pintar: " + totalpartidas);
                cuenta = 0;

                //creacion de las listas para productos y servicios y saber cuales estan inmpresos y cuales no
                List<string> rowids = new List<string>();
                List<string> rowidserv = new List<string>();
                int conteo = 0;
                int contadoarr = 0;

                query = "SELECT id FROM rowsipl WHERE ord ='" + idinvo.Text + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        rowids.Add(row["id"] + "|no");
                        conteo = conteo + 1;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();


                query = "SELECT id FROM rowsservpl WHERE ord ='" + idinvo.Text + "' ORDER BY id ASC;";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        rowidserv.Add(row["id"] + "|no");
                        conteo = conteo + 1;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();




                //arrfilas[conteo] = "no|" + row["id"];
                String[] arrfilas = rowids.ToArray();
                String[] arrfilasedit = rowids.ToArray();
                //FIN DE LA CREACION creacion de las listas para productos y servicios y saber cuales estan inmpresos y cuales no


                //FOR EACH MAESTRO DE PIEZAS Y SERVICIOS

                cajastot = 0;
                palletstot = 0;
                piezastot = 0;
                for (int i = 1; i <= totalpartidas; i++)
                {
                    /*
                    string logo = @"" + Path.GetDirectoryName(Application.ExecutablePath) + @"\" + "logo.jpg";
                    if (!File.Exists(@"" + logo))
                    {
                        //throw new FileNotFoundException(String.Format("No se encuentra el Logo {0}.", logo));
                    }
                    else
                    {
                        //XImage xImage = XImage.FromFile(logo);
                        //gfx.DrawImage(xImage, 20, 20, 104, 60);
                    }
                    */
                    //MessageBox.Show("cuantas van : " + cuantasvan);
                    if (cuantasvan == 0)
                    {

                        x = 0;
                        /*
                        gfx.DrawLine(pen, 1, 20, 1, 703);
                        gfx.DrawLine(pen, x + 50, 20, x + 50, 703);
                        gfx.DrawLine(pen, x + 100, 20, x + 100, 703);
                        gfx.DrawLine(pen, x + 150, 20, x + 150, 703);
                        gfx.DrawLine(pen, x + 200, 20, x + 200, 703);
                        gfx.DrawLine(pen, x + 250, 20, x + 250, 703);
                        gfx.DrawLine(pen, x + 300, 20, x + 300, 703);
                        gfx.DrawLine(pen, x + 350, 20, x + 350, 703);
                        gfx.DrawLine(pen, x + 400, 20, x + 400, 703);
                        gfx.DrawLine(pen, x + 450, 20, x + 450, 703);
                        gfx.DrawLine(pen, x + 500, 20, x + 500, 703);
                        gfx.DrawLine(pen, x + 550, 20, x + 550, 703);
                        gfx.DrawLine(pen, x + 600, 20, x + 600, 703);
                        gfx.DrawLine(pen, x + 650, 20, x + 650, 703);


                        gfx.DrawLine(pen, 45, 250, 45, 703);
                        gfx.DrawLine(pen, 87, 250, 87, 703);
                        gfx.DrawLine(pen, 150, 250, 150, 703);
                        gfx.DrawLine(pen, 291, 250, 291, 703);
                        gfx.DrawLine(pen, 381, 250, 381, 703);
                        gfx.DrawLine(pen, 461, 250, 461, 703);
                        gfx.DrawLine(pen, 571, 250, 571, 703);
                        */






                        //DATOS de encabezado
                        gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                        gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                        gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                        //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                        //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                        gfx.DrawString("PACKING LIST", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                        y = 85;

                        //DATOS de encabezado CLIENTE
                        gfx.DrawRectangle(pen, 13, 85, 330, 65);
                        gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                        //A LA DERECHA
                        gfx.DrawRectangle(pen, 400, 85, 180, 55);
                        gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PACKING LIST Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                        //ENCABEZADOS PARTIDAS
                        y = y + 70;
                        gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                        gfx.DrawString("DESCRIPCIÓN DE LA MERCANCÍA", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("UNIDADES", ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO BRUTO", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                        gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                        gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                        gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                        //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                        y = y + 20;
                        gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                        gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                        gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                        gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                        gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                        gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                        gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                        gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                        gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                        gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                        //gfx.DrawLine(pen, 400, y - 2, 400, y + 12);


                        // sub encabezados pequeños

                        // fin de sub encabezados pequeños
                        y = y + 15;

                    }

                    y = y + 13;
                    x = 100;
                    inity = 70;
                    cuantasvan = 1;
                    //gfx.DrawString(cuantasvan + " partidad:" + nife, SIETE, XBrushes.Black, new XRect(200, 49, 100, y), XStringFormats.Center);
                    //gfx.DrawString("Pag: " + conto + " de " + contgral, SIETE, XBrushes.Black, new XRect(500, 25, 100, 25), XStringFormats.Center);

                    #endregion headerproformpl

                    #region detailspl


                    //PRODUCTOS

                    //Un elemento de la lista puede cambiar su valor de manera similar usando el índice combinado con el operador de asignación.
                    //Por ejemplo, para cambiar el color de verde a mamey:
                    //ListaColores[2] = "mamey";




                    //PRODUCTOS

                    int cuentaprod = rowids.Count();

                    /*
                    contadoarr = 0;
                    foreach (string dato in arrfilas)
                    {
                        MessageBox.Show(i + ") Ver contenido del array editable: " + arrfilasedit[contadoarr]);
                        arrfilasedit[contadoarr] = "" + dato.Replace("no","si");
                        contadoarr = contadoarr + 1;
                    }
                    contadoarr = 0;

                    foreach (string dato in arrfilas)
                    {
                        MessageBox.Show(i + ") Ver contenido del array editable: " + arrfilasedit[contadoarr]);
                        contadoarr = contadoarr + 1;
                    }
                    */
                    contadoarr = 0;

                    query = "SELECT * FROM rowsipl WHERE ord ='" + idinvo.Text + "' ORDER BY container,id ASC;";
                    cm = new SqlCommand(query, con);
                    da = new SqlDataAdapter(cm);
                    dt = new DataTable();
                    da.Fill(dt);
                    cuenta = dt.Rows.Count;
                    if (cuenta > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            foreach (string dato in arrfilas)
                            {

                                if (contadoarr <= cuentaprod - 1)
                                {
                                    //MessageBox.Show( contadoarr +"Buscando porque no recorre ?: " + row["id"].ToString() + " array: " + arrfilasedit[contadoarr]);
                                    //MessageBox.Show(i +") Antes de eval: " + arrfilasedit[contadoarr]);
                                    //string[] datos = arrfilasedit[contadoarr].Split(new char[] { '|' });
                                    //Si ocupa el segundo es el índice 1 pues el primero es índice 0.
                                    //string idero = datos[0];
                                    //string yesno = datos[1];
                                    //MessageBox.Show("Buscando: " + arrfilas[contadoarr] +  " contra: " + idero +" que debe ser igual a: " + row["id"].ToString());
                                    // MessageBox.Show("estado actual: " + dato);

                                    //MessageBox.Show("Comparando: " + row["id"].ToString() + "|no"  + " " + arrfilasedit[contadoarr] );
                                    if (row["id"].ToString() + "|no" == arrfilasedit[contadoarr])
                                    {
                                        arrfilasedit[contadoarr] = "" + dato.Replace("no", "si");
                                        //  arrfilasedit[contadoarr] = "" + row["id"].ToString() + "|si";
                                        //MessageBox.Show(" -se imprimio el " + row["id"].ToString() + " - debe decir yes: "  + arrfilasedit[contadoarr]);

                                        string querya = "SELECT * FROM artsipl WHERE clave ='" + row["clave"] + "';";
                                        SqlCommand cma = new SqlCommand(querya, con);
                                        SqlDataAdapter daa = new SqlDataAdapter(cma);
                                        DataTable dta = new DataTable();
                                        daa.Fill(dta);
                                        int cuentaa = dta.Rows.Count;
                                        if (cuentaa > 0)
                                        {
                                            foreach (DataRow rowa in dta.Rows)
                                            {
                                                size = "" + rowa["size"]; sizel = "" + rowa["sizel"];
                                                caja = "" + rowa["caja"]; mtscaja = "" + rowa["mtscaja"]; kgspiece = "" + rowa["kgspiece"];
                                                kgscaja = "" + rowa["kgscaja"]; sku = "" + rowa["size"];
                                                pallet = "" + rowa["pallet"];
                                                units = "" + rowa["ume"];
                                            }

                                        }
                                        cma.Dispose(); daa.Dispose(); dta.Dispose();



                                        ctns = 0;
                                        sqm = 0;
                                        tctns = 0;
                                        tsqm = 0;
                                        cajad = 0;
                                        mtsd = 0;
                                        pallets = 0;
                                        pieces = 0;

                                        try { ctns = double.Parse("" + caja); }
                                        catch { ctns = 0; }

                                        try { sqm = double.Parse("" + mtscaja); }
                                        catch { sqm = 0; }

                                        try { pallets = double.Parse("" + pallet); }
                                        catch { pallets = 0; }


                                        try
                                        {
                                            pieces = (ctns * double.Parse("" + row["cant"])) / sqm;
                                        }
                                        catch
                                        {
                                            pieces = 0;
                                        }


                                        try
                                        {
                                            tctns = double.Parse("" + row["cant"]) / sqm;
                                        }
                                        catch { tctns = 0; }

                                        try
                                        {
                                            tsqm = sqm * tctns;
                                        }
                                        catch { tsqm = 0; }
                                        if (row["pallets"].ToString() != "")
                                        {
                                            pallets = double.Parse("" + row["pallets"]);
                                        }
                                        else
                                        {
                                            try
                                            {
                                                pallets = double.Parse("" + row["cant"]) / pallets;
                                            }
                                            catch { pallets = 1; }
                                            //if (pallets < 1) { pallets = 1; }
                                            if (pallets.ToString() == "∞") { pallets = 0; }
                                        }
                                        try
                                        {
                                            cajad = double.Parse("" + row["cant"]) / sqm;
                                        }
                                        catch { cajad = 0; }



                                        if (contadoarr < cuentaprod - 1)
                                        {
                                            cajastot = cajastot + tctns;
                                            palletstot = palletstot + pallets;
                                            piezastot = piezastot + pieces;
                                        }
                                        //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                                        //gfx.DrawString("" + size +"X"+ sizel, ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString(""+ row["clave"].ToString(), ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("1", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        //gfx.DrawString("" + row["pallets"], ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);
                                        tf.Alignment = XParagraphAlignment.Right;
                                        gfx.DrawString("" + size + "X" + sizel, ocho, XBrushes.Black, new XRect(20, y, 100, 25), XStringFormats.TopLeft);
                                        gfx.DrawString("" + row["clave"].ToString(), SIETE, XBrushes.Black, new XRect(60, y, 100, 25), XStringFormats.TopLeft);
                                        gfx.DrawString("1", ocho, XBrushes.Black, new XRect(163, y, 100, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + pallets.ToString("n2"), ocho, XBrushes.Black, new XRect(175, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + Convert.ToInt32(tctns).ToString("N0"), ocho, XBrushes.Black, new XRect(235, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(295, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + Convert.ToInt32(pieces).ToString("N0"), ocho, XBrushes.Black, new XRect(340, y, 50, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + row["pesoneto"], ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                                        tf.DrawString("" + row["pesobruto"], ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);

                                        y = y + 13;



                                        // INIT bloque para poder imprimir el totalizador de container
                                        string bsql = "SELECT TOP (1) * FROM rowsipl WHERE ord = '" + row["ord"].ToString() + "' AND container= '" + row["container"].ToString() + "' ORDER BY id desc";
                                        SqlCommand cmab = new SqlCommand(bsql, con);
                                        SqlDataAdapter daab = new SqlDataAdapter(cmab);
                                        DataTable dtab = new DataTable();
                                        daab.Fill(dtab);
                                        int cuentaab = dtab.Rows.Count;

                                        if (cuentaab > 0)
                                        {
                                            foreach (DataRow rowaa in dtab.Rows)
                                            {
                                                if (rowaa["id"].ToString() + "" == "" + row["id"].ToString())
                                                {
                                                    string bsqlc = "SELECT precinto,pesoneto,pesobruto,container FROM containersipl WHERE ord = '" + row["ord"].ToString() + "' AND container= '" + row["container"].ToString() + "' ORDER BY id desc";
                                                    SqlCommand cmabc = new SqlCommand(bsqlc, con);
                                                    SqlDataAdapter daabc = new SqlDataAdapter(cmabc);
                                                    DataTable dtabc = new DataTable();
                                                    daabc.Fill(dtabc);
                                                    int cuentaabc = dtabc.Rows.Count;

                                                    if (cuentaabc > 0)
                                                    {
                                                        foreach (DataRow rowaac in dtabc.Rows)
                                                        {
                                                            //CONTENEDOR: RNHC6213642 PRECINTO: SIN PRECINTO              48,435.36   49,259.76

                                                            //MessageBox.Show("Imprimir remate totalizador del container: " + rowaa["id"].ToString() + " " + row["container"].ToString());


                                                            gfx.DrawRectangle(peng, 13, y - 2, 580, 14);
                                                            gfx.DrawLine(peng, 190, y - 2, 190, y + 12);
                                                            gfx.DrawLine(peng, 400, y - 2, 400, y + 12);
                                                            gfx.DrawLine(peng, 490, y - 2, 490, y + 12);


                                                            tf.Alignment = XParagraphAlignment.Right;
                                                            gfx.DrawString("CONTAINER: " + rowaac["container"].ToString(), ocho, XBrushes.Black, new XRect(20, y, 100, 25), XStringFormats.TopLeft);
                                                            //gfx.DrawString("", SIETE, XBrushes.Black, new XRect(60, y, 100, 25), XStringFormats.TopLeft);
                                                            gfx.DrawString("", ocho, XBrushes.Black, new XRect(163, y, 100, 25), XStringFormats.TopLeft);
                                                            tf.DrawString("PRECINTO: " + rowaac["precinto"].ToString(), ocho, XBrushes.Black, new XRect(100, y, 200, 25), XStringFormats.TopLeft);
                                                            //tf.DrawString("" + Convert.ToInt32(tctns).ToString("N0"), ocho, XBrushes.Black, new XRect(235, y, 50, 25), XStringFormats.TopLeft);
                                                            //tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(295, y, 50, 25), XStringFormats.TopLeft);
                                                            //tf.DrawString("" + Convert.ToInt32(pieces).ToString("N0"), ocho, XBrushes.Black, new XRect(340, y, 50, 25), XStringFormats.TopLeft);
                                                            tf.DrawString("" + rowaac["pesoneto"].ToString(), ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                                                            tf.DrawString("" + rowaac["pesobruto"].ToString(), ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);
                                                        }

                                                    }

                                                    y = y + 22;
                                                    cuantasvan = cuantasvan + 1;
                                                }
                                            }
                                        }
                                        // FIN bloque para poder imprimir el totalizador de container


                                        cuantasvan = cuantasvan + 1;
                                        cuantastotoal = cuantastotoal + 1;
                                        //SECCION PARA PODER AGREGAR PAGINAS NUEVAS AL PDF
                                        if (cuantasvan == 42)
                                        {
                                            if (cuantastotoal == totalpartidas)
                                            {

                                            }
                                            else
                                            {
                                                page = document.AddPage();
                                                page.Orientation = PdfSharp.PageOrientation.Portrait;
                                                page.Size = PdfSharp.PageSize.Letter;
                                                gfx = XGraphics.FromPdfPage(page);
                                                tf = new XTextFormatter(gfx);
                                                cuantasvan = 1;
                                                //DATOS de encabezado
                                                gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                                                //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                                                //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                                                gfx.DrawString("PACKING LIST", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                                                y = 85;

                                                //DATOS de encabezado CLIENTE
                                                gfx.DrawRectangle(pen, 13, 85, 330, 65);
                                                gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                                                //A LA DERECHA
                                                gfx.DrawRectangle(pen, 400, 85, 180, 55);
                                                gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("FACTURA Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                                                //ENCABEZADOS PARTIDAS
                                                y = y + 70;
                                                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                                                gfx.DrawString("DESCRIPCIÓN DE LA MERCANCÍA", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("UNIDADES", ocho, XBrushes.Black, new XRect(265, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO BRUTO", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                                                gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                                                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                                                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                                                //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                                                y = y + 20;
                                                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                                                gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                                                gfx.DrawString("PESO", ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                                                gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                                                gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                                                gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                                                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                                                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                                                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                                                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                                                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                                                y = y + 15;
                                                y = y + 13;
                                                x = 100;
                                                inity = 70;
                                                cuantasvan = 1;
                                            }
                                        }
                                        //FIN DE SECCION PARA PODER AGREGAR PAGINAS NUEVAS AL PDF
                                    }
                                    contadoarr = contadoarr + 1;
                                }
                            }     // for each que recorre el array maestro
                            contadoarr = 0;
                        }// for each que recorre el bloque de productos en la factura desde la base
                    }
                    cm.Dispose(); da.Dispose(); dt.Dispose();
                    contadoarr = 0;
                    #endregion detailspl

                    #region footerpl
                } // ESTE ES EL FOR EACH MAESTRO DE LOS REGISTROS DE PIEZAS Y SERVICIOS

                //HOJA ESPECIAL DE RESUMEN
                page = document.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Portrait;
                page.Size = PdfSharp.PageSize.Letter;
                gfx = XGraphics.FromPdfPage(page);
                tf = new XTextFormatter(gfx);

                cuantasvan = 1;
                //DATOS de encabezado
                gfx.DrawString("" + nome, dieznegra, XBrushes.Black, new XRect(250, 18, 100, 25), XStringFormats.Center);
                gfx.DrawString("" + "" + callee + " " + nume + " " + numie + " " + cole + " " + cde + " " + estadoe + " " + paise + " CP: " + " " + cpe, SIETE, XBrushes.Black, new XRect(250, 28, 100, 25), XStringFormats.Center);
                gfx.DrawString("" + "CIF: " + nife, SIETE, XBrushes.Black, new XRect(250, 38, 100, 25), XStringFormats.Center);
                //gfx.DrawString("" + e5 + "    C.P.: " + e7, SIETE, XBrushes.Black, new XRect(200, 35, 100, 25), XStringFormats.Center);
                //gfx.DrawString("" + e6 + "    TEL.:" + e8, SIETE, XBrushes.Black, new XRect(200, 42, 100, 25), XStringFormats.Center);
                gfx.DrawString("PACKING LIST - RESUMEN", dieznegra, XBrushes.Black, new XRect(250, 49, 100, 35), XStringFormats.Center);

                y = 85;

                //DATOS de encabezado CLIENTE
                gfx.DrawRectangle(pen, 13, 85, 330, 65);
                gfx.DrawString("" + nomcli, ochoneg, XBrushes.Black, new XRect(16, y + 3, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + nifcli, SIETE, XBrushes.Black, new XRect(16, y + 12, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + callecli + " " + numcli + " " + numclii, SIETE, XBrushes.Black, new XRect(16, y + 22, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + colcli, SIETE, XBrushes.Black, new XRect(16, y + 30, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + muncli + ", " + edocli, SIETE, XBrushes.Black, new XRect(16, y + 43, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + paiscli + ", CP:" + cpcli, SIETE, XBrushes.Black, new XRect(16, y + 56, 100, 25), XStringFormats.TopLeft);

                //A LA DERECHA
                gfx.DrawRectangle(pen, 400, 85, 180, 55);
                gfx.DrawString("FECHA: " + fecha, ocho, XBrushes.Black, new XRect(404, y + 3, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("Nº: " + folio + "", ocho, XBrushes.Black, new XRect(404, y + 14, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("FACTURA Nº: " + numees, ocho, XBrushes.Black, new XRect(404, y + 27, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("PARTIDA ESTADISTICA: " + albaran, ocho, XBrushes.Black, new XRect(404, y + 40, 100, 25), XStringFormats.TopLeft);


                //ENCABEZADOS PARTIDAS
                y = y + 70;
                /*
                gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                gfx.DrawString("GASTOS INDIRECTOS", ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("CANTIDAD", ocho, XBrushes.Black, new XRect(355, y, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("PRECIO NETO", ocho, XBrushes.Black, new XRect(420, y, 100, 25), XStringFormats.TopLeft);
                gfx.DrawString("IMPORTE TOTAL", ocho, XBrushes.Black, new XRect(505, y, 50, 25), XStringFormats.TopLeft);

                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);

                */

                //FORMATO	MODELO	CLASE	PALLETS	CAJAS	M²	PIEZAS	USD	USD
                /*
                y = y + 20;
                    gfx.DrawRectangle(pen, 13, y - 2, 580, 14);
                    gfx.DrawString("FORMATO", SIETE, XBrushes.Black, new XRect(16, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("MODELO", ocho, XBrushes.Black, new XRect(72, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("CLASE", ocho, XBrushes.Black, new XRect(156, y, 100, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                    gfx.DrawString("" + currency, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);


                    gfx.DrawLine(pen, 53, y - 2, 53, y + 12);
                    gfx.DrawLine(pen, 147, y - 2, 147, y + 12);
                    gfx.DrawLine(pen, 190, y - 2, 190, y + 12);
                    gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                    gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                    gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                    gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                    gfx.DrawLine(pen, 490, y - 2, 490, y + 12);
                    */

                y = y + 13;
                // SERVICIOS
                /*
                query = "SELECT *  FROM rowsservpl WHERE ord ='" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        tf.Alignment = XParagraphAlignment.Right;
                        gfx.DrawString("" + row["descrip"], ocho, XBrushes.Black, new XRect(35, y, 100, 25), XStringFormats.TopLeft);
                        tf.DrawString("" + row["cant"], ocho, XBrushes.Black, new XRect(298, y, 100, 25), XStringFormats.TopLeft);
                        tf.DrawString(csymbol_l + "" + "" + row["cu"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(400, y, 80, 25), XStringFormats.TopLeft);
                        tf.DrawString(csymbol_l + "" + row["total"] + "" + csymbol_r, ocho, XBrushes.Black, new XRect(505, y, 80, 25), XStringFormats.TopLeft);
                        y = y + 13;
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
                */
                y = y + 15;
                y = y + 13;
                x = 100;
                inity = 70;

                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("PALLETS", ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("CAJAS", ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("M²", ocho, XBrushes.Black, new XRect(320, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PIEZAS", ocho, XBrushes.Black, new XRect(358, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO NETO", ocho, XBrushes.Black, new XRect(415, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO BRUTO", ocho, XBrushes.Black, new XRect(510, y, 50, 25), XStringFormats.TopLeft);



                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);



                double metrostotals = 0;
                query = "SELECT SUM(convert(numeric(18, 6), replace(cant, ',', ''))) as totals FROM rowsipl  where ord = '" + idinvo.Text + "';";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        metrostotals = double.Parse("" + row["totals"]);
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();

                cajastot = cajastot + tctns; palletstot = palletstot + pallets; piezastot = piezastot + pieces;

                y = y + 20;
                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("" + palletstot.ToString("N2"), ocho, XBrushes.Black, new XRect(198, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + cajastot.ToString("N0"), ocho, XBrushes.Black, new XRect(253, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + metrostotals.ToString("N2"), ocho, XBrushes.Black, new XRect(305, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + piezastot.ToString("N0"), ocho, XBrushes.Black, new XRect(357, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + netofact, ocho, XBrushes.Black, new XRect(430, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + brutosfact, ocho, XBrushes.Black, new XRect(535, y, 50, 25), XStringFormats.TopLeft);



                gfx.DrawLine(pen, 240, y - 2, 240, y + 12);
                gfx.DrawLine(pen, 295, y - 2, 295, y + 12);
                gfx.DrawLine(pen, 350, y - 2, 350, y + 12);
                gfx.DrawLine(pen, 400, y - 2, 400, y + 12);
                gfx.DrawLine(pen, 490, y - 2, 490, y + 12);


                y = y + 60;
                gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                gfx.DrawString("CONTENEDOR(ES)", ocho, XBrushes.Black, new XRect(197, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PRECINTO", ocho, XBrushes.Black, new XRect(290, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO NETO EN KGS", ocho, XBrushes.Black, new XRect(400, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("PESO BRUTO EN KGS", ocho, XBrushes.Black, new XRect(500, y, 50, 25), XStringFormats.TopLeft);

                gfx.DrawLine(pen, 288, y - 2, 288, y + 12);
                gfx.DrawLine(pen, 396, y - 2, 396, y + 12);
                gfx.DrawLine(pen, 498, y - 2, 498, y + 12);



                y = y + 17;
                string queryc = "SELECT * FROM containersipl WHERE ord ='" + id + "' ORDER BY id ASC;";
                SqlCommand cmc = new SqlCommand(queryc, con);
                SqlDataAdapter dac = new SqlDataAdapter(cmc);
                DataTable dtc = new DataTable();
                dac.Fill(dtc);
                int cuentac = dtc.Rows.Count;
                if (cuentac > 0)
                {
                    foreach (DataRow rowc in dtc.Rows)
                    {
                        gfx.DrawRectangle(pen, 195, y - 2, 400, 14);
                        gfx.DrawString("" + rowc["container"], ocho, XBrushes.Black, new XRect(197, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + rowc["precinto"], ocho, XBrushes.Black, new XRect(293, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + double.Parse("" + rowc["pesoneto"]), ocho, XBrushes.Black, new XRect(403, y, 50, 25), XStringFormats.TopLeft);
                        gfx.DrawString("" + double.Parse("" + rowc["pesobruto"]), ocho, XBrushes.Black, new XRect(503, y, 50, 25), XStringFormats.TopLeft);

                        gfx.DrawLine(pen, 288, y - 2, 288, y + 12);
                        gfx.DrawLine(pen, 396, y - 2, 396, y + 12);
                        gfx.DrawLine(pen, 498, y - 2, 498, y + 12);

                        y = y + 17;
                    }
                }
                cmc.Dispose(); dac.Dispose(); dtc.Dispose();



                y = y + 20;
                gfx.DrawRectangle(pen, 30, y - 2, 565, 14);
                gfx.DrawString("DESTINO: ", ocho, XBrushes.Black, new XRect(33, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawString("" + origdest, ocho, XBrushes.Black, new XRect(85, y, 50, 25), XStringFormats.TopLeft);
                gfx.DrawLine(pen, 80, y - 2, 80, y + 12);

                y = y + 25;
                /*
                tf.Alignment = XParagraphAlignment.Right;
                gfx.DrawRectangle(pen, 376, y - 2, 220, 14);
                gfx.DrawString("TOTAL FACTURA", dieznegra, XBrushes.Black, new XRect(380, y, 50, 25), XStringFormats.TopLeft);
                tf.DrawString("" + csymbol_l + "" + tot + "" + csymbol_r, dieznegra, XBrushes.Black, new XRect(500, y, 90, 25), XStringFormats.TopLeft);
                gfx.DrawLine(pen, 478, y - 2, 478, y + 12);
                */

                y = y + 25;

                tf.Alignment = XParagraphAlignment.Justify;
                gfx.DrawRectangle(pen, 30, y - 2, 565, 120);
                gfx.DrawString("Observaciones: ", ocho, XBrushes.Black, new XRect(33, y, 50, 25), XStringFormats.TopLeft);
                y = y + 13;
                tf.DrawString("" + obs1 + "\r\n" + obs2 + "\r\n" + obs3 + "\r\n" + obs4 + "\r\n" + obs5, CINCO, XBrushes.Black, new XRect(40, y, 548, 100), XStringFormats.TopLeft);



                #endregion footerpl

                #region openerpl

                try
                {
                    document.Save(@"" + paths);
                    // ...and start a viewer.
                    Process.Start(paths);
                }
                catch (Exception ex) { MessageBox.Show("Actualizando, presione el boton de nuevo\r\n" + ex.ToString()); }
                #endregion openerpl
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            #region init

            string currency="", csymbol = "";

            int rows = 2, rowed = 2;
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            DateTime hoy = DateTime.Now;
            string dates = hoy.Day.ToString() + "-" + hoy.Month.ToString() + "-" + hoy.Year.ToString() + "-" + hoy.Hour.ToString() + "-" + hoy.Minute.ToString() + "-" + hoy.Second.ToString();
            ExcelPackage excel = new ExcelPackage();
            excel.Workbook.Worksheets.Add("RESUMEN");
            
            try
            {
                Directory.EnumerateFiles(@"" + config.tempofiles, "RESUMEN_*.xlsx").ToList().ForEach(x => File.Delete(x));
            }
            catch { }
            FileInfo excelFile = new FileInfo(@"" + config.tempofiles + @"\RESUMEN_" + dates + ".xlsx");
            var proforma = excel.Workbook.Worksheets["RESUMEN"];
            SqlConnection con = new SqlConnection("" + config.cade);
            con.Open();
            #endregion init

            #region Headers


            /*
folio
num
status
nombre cliente
pais
fecha
albaran
origen / destino
moneda
Total
*/


            proforma.Cells["A1:L1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            proforma.Cells["A1:L1"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


            proforma.Cells["A1"].Style.Font.Bold = true;
            proforma.Cells["A1"].Style.Font.Size = 11;
            proforma.Cells["A1"].Style.Font.Name = "Calibri";
            proforma.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            proforma.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["A1"].Value = "FOLIO";

            proforma.Cells["B1"].Style.Font.Bold = true;
            proforma.Cells["B1"].Style.Font.Size = 11;
            proforma.Cells["B1"].Style.Font.Name = "Calibri";
            proforma.Cells["B1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            proforma.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["B1"].Value = "NUM";


            proforma.Cells["C1"].Style.Font.Bold = true;
            proforma.Cells["C1"].Style.Font.Size = 11;
            proforma.Cells["C1"].Style.Font.Name = "Calibri";
            proforma.Cells["C1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            proforma.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["C1"].Value = "STATUS";

            proforma.Cells["D1"].Style.Font.Bold = true;
            proforma.Cells["D1"].Style.Font.Size = 11;
            proforma.Cells["D1"].Style.Font.Name = "Calibri";
            proforma.Cells["D1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            proforma.Cells["D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["D1"].Value = "No.Cli";

            proforma.Cells["E1"].Style.Font.Bold = true;
            proforma.Cells["E1"].Style.Font.Size = 11;
            proforma.Cells["E1"].Style.Font.Name = "Calibri";
            proforma.Cells["E1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            proforma.Cells["E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["E1"].Value = "CLIENTE";

            proforma.Cells["F1"].Style.Font.Bold = true;
            proforma.Cells["F1"].Style.Font.Size = 11;
            proforma.Cells["F1"].Style.Font.Name = "Calibri";
            proforma.Cells["F1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            proforma.Cells["F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["F1"].Value = "PAIS";

            proforma.Cells["G1"].Style.Font.Bold = true;
            proforma.Cells["G1"].Style.Font.Size = 11;
            proforma.Cells["G1"].Style.Font.Name = "Calibri";
            proforma.Cells["G1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            proforma.Cells["G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["G1"].Value = "FECHA";

            proforma.Cells["H1"].Style.Font.Bold = true;
            proforma.Cells["H1"].Style.Font.Size = 11;
            proforma.Cells["H1"].Style.Font.Name = "Calibri";
            proforma.Cells["H1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            proforma.Cells["H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["H1"].Value = "ALBARAN";


            proforma.Cells["I1"].Style.Font.Bold = true;
            proforma.Cells["I1"].Style.Font.Size = 11;
            proforma.Cells["I1"].Style.Font.Name = "Calibri";
            proforma.Cells["I1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            proforma.Cells["I1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["I1"].Value = "ORIGEN/DESTINO";

            proforma.Cells["J1"].Style.Font.Bold = true;
            proforma.Cells["J1"].Style.Font.Size = 11;
            proforma.Cells["J1"].Style.Font.Name = "Calibri";
            proforma.Cells["J1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            proforma.Cells["J1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["J1"].Value = "MONEDA";


            proforma.Cells["K1"].Style.Font.Bold = true;
            proforma.Cells["K1"].Style.Font.Size = 11;
            proforma.Cells["K1"].Style.Font.Name = "Calibri";
            proforma.Cells["K1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            proforma.Cells["K1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            proforma.Cells["K1"].Value = "TOTAL";

            proforma.Cells["L1"].Style.Font.Bold = true;
            proforma.Cells["L1"].Style.Font.Size = 11;
            proforma.Cells["L1"].Style.Font.Name = "Calibri";
            proforma.Cells["L1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            proforma.Cells["L1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            #endregion Headers


            #region detail



            string inicial = "" + init.Text;
            string final = "" + finit.Text;

            var primera = Convert.ToDateTime(inicial);
            var segunda = Convert.ToDateTime(final);
            inicial = "" + primera.ToString("yyyy-MM-dd");
            final = "" + segunda.ToString("yyyy-MM-dd");
            string range = "";

            string cli = "" + numc.Text;
            string emi = "" + nume.Text;
            double totalo = 0;

            if (cli != "") { cli = " AND  idcli='" + cli + "' "; }
            if (emi != "") { emi = " AND  empresa='" + emi + "' "; }
            range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE) " + cli + " " + emi + ";";
            string sqlSelectAll = invoice_query + " " + range + ";";



            string query = "" + sqlSelectAll;
            SqlCommand cm = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(cm);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    currency = "" + row["Moneda"];

                    if (currency == "USD") { csymbol = "\"US$\"#,##0.00;-\"US$\"#,##0.00"; }
                    if (currency == "EUR") { csymbol = "\"\"#,##0.00€;-\"\"#,##0.00€"; }
                    if (currency == "") { csymbol = "\"\"#,##0.00;-\"\"#,##0.00"; }

                    //"SELECT folio AS FOLIO, number AS NUM,empresa AS #Emp, idcli AS #Cli,nomcli AS Nombre, paiscli AS Pais, fecha as Fecha, albaran as Albaran, origdest as [Origen Destino],tot AS Total, currency as Moneda ,id,stats AS [Status] FROM invoicespl"
                    proforma.Cells["A"+ rows].Style.Font.Bold = true;
                    proforma.Cells["A" + rows].Style.Font.Size = 9;
                    proforma.Cells["A" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["A" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["A" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["A" + rows].Value = "" + row["FOLIO"];

                    proforma.Cells["B" + rows].Style.Font.Bold = true;
                    proforma.Cells["B" + rows].Style.Font.Size = 9;
                    proforma.Cells["B" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["B" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["B" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["B" + rows].Value = "" + row["NUM"];

                    proforma.Cells["C" + rows].Style.Font.Bold = true;
                    proforma.Cells["C" + rows].Style.Font.Size = 9;
                    proforma.Cells["C" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["C" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["C" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["C" + rows].Value = "" + row["Status"];

                    proforma.Cells["D" + rows].Style.Font.Bold = true;
                    proforma.Cells["D" + rows].Style.Font.Size = 9;
                    proforma.Cells["D" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["D" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["D" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["D" + rows].Value = "" + row["#Cli"];

                    proforma.Cells["E" + rows].Style.Font.Bold = true;
                    proforma.Cells["E" + rows].Style.Font.Size = 9;
                    proforma.Cells["E" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["E" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    proforma.Cells["E" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    proforma.Cells["E" + rows].Value = "" + row["Nombre"];

                    proforma.Cells["F" + rows].Style.Font.Bold = true;
                    proforma.Cells["F" + rows].Style.Font.Size = 9;
                    proforma.Cells["F" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["F" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["F" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["F" + rows].Value = "" + row["Pais"];

                    proforma.Cells["G" + rows].Style.Font.Bold = true;
                    proforma.Cells["G" + rows].Style.Font.Size = 9;
                    proforma.Cells["G" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["G" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["G" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["G" + rows].Value = "" + row["Fecha"];

                    proforma.Cells["H" + rows].Style.Font.Bold = true;
                    proforma.Cells["H" + rows].Style.Font.Size = 9;
                    proforma.Cells["H" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["H" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["H" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["H" + rows].Value = "" + row["Albaran"];

                    proforma.Cells["I" + rows].Style.Font.Bold = true;
                    proforma.Cells["I" + rows].Style.Font.Size = 9;
                    proforma.Cells["I" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["I" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    proforma.Cells["I" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    proforma.Cells["I" + rows].Value = "" + row["Origen Destino"];

                    proforma.Cells["J" + rows].Style.Font.Bold = true;
                    proforma.Cells["J" + rows].Style.Font.Size = 9;
                    proforma.Cells["J" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["J" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    proforma.Cells["J" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    proforma.Cells["J" + rows].Value = "" + row["Moneda"];


                    totalo = double.Parse("" + row["Total"]);
                    proforma.Cells["K" + rows].Style.Font.Bold = true;
                    proforma.Cells["K" + rows].Style.Font.Size = 9;
                    proforma.Cells["K" + rows].Style.Font.Name = "Calibri";
                    proforma.Cells["K" + rows].Style.Numberformat.Format =  csymbol;
                    proforma.Cells["K" + rows].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    proforma.Cells["K" + rows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    proforma.Cells["K" + rows].Value = totalo;

                    rows = rows + 1;
                    rowed = rows;
                }

            }
            cm.Dispose();dt.Dispose();da.Dispose();

            #endregion detail

            #region opener

            rowed = rowed - 1;
            proforma.Cells["L1"].Formula = "=SUBTOTAL(9,K2:K" + rowed + ")";
            proforma.Cells["L1"].Style.Numberformat.Format = "#,##0.00";
           

            proforma.View.FreezePanes(2, 1);
            proforma.Cells["A1:K1"].AutoFilter = true;
            proforma.Cells["A1:K1"].AutoFitColumns();
            
            proforma.Column(1).AutoFit();
            proforma.Column(1).Width = 13;

            proforma.Column(2).AutoFit();
            proforma.Column(2).Width = 13;
            proforma.Column(3).AutoFit();
            proforma.Column(3).Width =12;
            proforma.Column(4).Width = 8;
            proforma.Column(5).Width = 35;
            proforma.Column(6).Width = 8;
            proforma.Column(7).AutoFit();
            proforma.Column(7).Width = 9;
            proforma.Column(8).Width = 13;
            proforma.Column(9).Width = 20;
            proforma.Column(10).Width = 11;
            proforma.Column(11).Width = 15;
            proforma.Column(12).Width = 16;
            //proforma.Column(8).AutoFit();


            double TOPMA = 0;
            double LEFTMA = 0;
            try
            {

                TOPMA = double.Parse(supo.Text);
            }
            catch { TOPMA = 10; }
            try
            {
                LEFTMA = double.Parse(lefte.Text);
            }
            catch { LEFTMA = 10; }

            try { TOPMA = TOPMA / 10; }
            catch { TOPMA = 1; }

            try { LEFTMA = LEFTMA / 10; }
            catch { LEFTMA = 1; }

            proforma.PrinterSettings.TopMargin = (decimal)TOPMA / 2.54M; // narrow border
            proforma.PrinterSettings.RightMargin = (decimal).4 / 2.54M; //narrow border
            proforma.PrinterSettings.LeftMargin = (decimal)LEFTMA / 2.54M; //narrow border
            proforma.PrinterSettings.BottomMargin = (decimal).4 / 2.54M; //narrow border

            //proforma.Row(30).PageBreak = true;
            proforma.PrinterSettings.PaperSize = ePaperSize.Letter;
            proforma.PrinterSettings.Orientation = eOrientation.Portrait;
            //proforma.PrinterSettings.Scale = 75;

            excel.SaveAs(excelFile);
            System.Diagnostics.Process.Start(@"" + excelFile);
            con.Close();
            #endregion opener


        }

        private void deces_ValueChanged(object sender, EventArgs e)
        {
            deci = Int32.Parse("" + deces.Value);
        }
    }
}
