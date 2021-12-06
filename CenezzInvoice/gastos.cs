using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
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

namespace CenezzInvoice
{
    public partial class gastos : Form
    {
        public gastos()
        {
            InitializeComponent();
        }
        internal static string invoice_query = "SELECT clave AS Clave, descr AS [Descripión],FORMAT( convert(numeric(18,5),replace(precio,',','')),'###,###,###.00000','ES-mx') AS [Precio], id FROM gastos ORDER BY clave ASC ";
        internal static string invoice_querys = "SELECT clave AS Clave, descr AS [Descripión], FORMAT( convert(numeric(18,5),replace(precio,',','')),'###,###,###.00000','ES-mx') AS [Precio], id FROM gastos ";
        //internal static string invoice_query = "SELECT clave AS Clave, descr AS [Descripión], precio AS [Precio], id FROM gasto ORDER BY clave ASC ";

        private void gastos_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }
        private void gastos_Load(object sender, EventArgs e)
        {
            SqlConnection con = new
            SqlConnection("" + config.cade);
            con.Open();

            string range = "";
            //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
            SqlDataAdapter DA = new SqlDataAdapter();
            string sqlSelectAll = invoice_query + " " + range + ";";
            DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

            DataTable table = new DataTable();
            DA.Fill(table);

            BindingSource bSource = new BindingSource();
            bSource.DataSource = table;
            lister.DataSource = bSource;

            resizegrid();
            con.Close();

            SqlCommand cmd = new SqlCommand("SELECT clave FROM gastos ORDER BY clave ASC;", con);
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            AutoCompleteStringCollection MyCollection = new AutoCompleteStringCollection();
            while (reader.Read())
            {
                MyCollection.Add(reader.GetString(0));
            }
            cves.AutoCompleteCustomSource = MyCollection;
            cmd.Dispose(); reader.Dispose();


        }

        private void resizegrid()
        {
            lister.Columns[0].Width = 100;
            lister.Columns[1].Width = 600;
            lister.Columns[2].Width = 80;
            lister.Columns[3].Width = 50;
            lister.Columns[3].Visible = false;
        }


        private void gastos_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }



        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (idinvo.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea ELIMINAR?", "--Eliminar Servicio --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    string cellValueid = idinvo.Text;
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "DELETE FROM gastos WHERE id=" + cellValueid + ";";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    string range = "";
                    //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    string sqlSelectAll = invoice_query + " " + range + ";";
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = table;
                    lister.DataSource = bSource;

                    con.Close();

                }
            }
            else { MessageBox.Show("Debes seleccionar un Producto para eliminarlo."); }
        }

        private void lister_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            idinvo.Text = "";
            cvee.Text = "";
            descre.Text = "";
            precioe.Text = "";


            if (e.RowIndex != -1)
            {

                var dataIndexNo = lister.Rows[e.RowIndex].Index.ToString();
                string cellValue = lister.Rows[e.RowIndex].Cells[0].Value.ToString();
                string cellValueid = lister.Rows[e.RowIndex].Cells[3].Value.ToString();

                idinvo.Text = "" + cellValueid;
                SqlConnection con = new
                SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT * FROM gastos WHERE id=" + cellValueid + ";";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        cvee.Text = "" + row["clave"];
                        descre.Text = "" + row["descr"];
                      //  precioe.Text = "" + double.Parse("" + row["precio"]).ToString("n2");

                    }
                }
                da.Dispose(); cm.Dispose(); dt.Dispose();
                con.Close();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (cvee.Text != "" && idinvo.Text != "")
            {

                //precioe.Text = "" + double.Parse("" + precioe.Text).ToString("n2");
                if (DialogResult.Yes == MessageBox.Show("¿Desea guardar?", " -- Almacenar cambios del servicio --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "UPDATE gastos SET  descr='" + descre.Text + "', precio='" + precioe.Text + "' ";
                    qu = qu + " WHERE id=" + idinvo.Text + ";";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    string range = " WHERE id =" + idinvo.Text + ";";
                    //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    //string sqlSelectAll = invoice_query + " " + range + "";
                    string sqlSelectAll = "SELECT clave AS Clave, descr AS [Descripión], precio AS [Precio], id FROM gastos " + range;
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = table;
                    lister.DataSource = bSource;
                    con.Close();

                    cvee.Text = "";
                    descre.Text = "";
                    precioe.Text = "";
                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (cve.Text != "")
            {
                //precio.Text = "" + double.Parse("" + precio.Text).ToString("n2");
                if (DialogResult.Yes == MessageBox.Show("¿Desea guardar?", " -- Almacenar nuevo servicio --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "INSERT INTO gastos (clave, descr, precio) ";
                    qu = qu + "VALUES('" + cve.Text + "','" + descr.Text + "','" + precio.Text + "');SELECT SCOPE_IDENTITY();";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    string uuid = "" + myCo.ExecuteScalar().ToString();
                    myCo.Dispose();


                    string range = " WHERE id =" + uuid + ";";
                    //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    string sqlSelectAll = "SELECT clave AS Clave, descr AS [Descripión], precio AS [Precio], id FROM gastos " + range;
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = table;
                    lister.DataSource = bSource;
                   

                    cve.Text = "";
                    descr.Text = "";
                    precio.Text = "";

                    cvee.Text = "";
                    descre.Text = "";
                    precioe.Text = "";


                    SqlCommand cmd = new SqlCommand("SELECT clave FROM gastos ORDER BY clave ASC;", con);
                    SqlDataReader reader = cmd.ExecuteReader();
                    AutoCompleteStringCollection MyCollection = new AutoCompleteStringCollection();
                    while (reader.Read())
                    {
                        MyCollection.Add(reader.GetString(0));
                    }
                    cves.AutoCompleteCustomSource = MyCollection;
                    cmd.Dispose(); reader.Dispose();
                    con.Close();
                }
            }
            else
            {
                MessageBox.Show("No puedes dejar la clave vacia");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            SqlConnection con = new
            SqlConnection("" + config.cade);
            con.Open();

            if (cves.Text != "")
            {




                string range = "";
                string cvess = "", skuss = "", lineass = "";
                if (cves.Text != "")
                {
                    cvess = " AND clave LIKE '%" + cves.Text + "%' ";
                }
              

                range = "" + cvess + "" + skuss + "" + lineass;

                int largo = range.Length;
                if (largo >= 4)
                {
                    range = range.Substring(4);
                }

                range = " WHERE " + range;
                //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                SqlDataAdapter DA = new SqlDataAdapter();
                string sqlSelectAll = invoice_querys + " " + range + "  ORDER BY clave ASC ;";
                DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                DataTable table = new DataTable();
                DA.Fill(table);

                BindingSource bSource = new BindingSource();
                bSource.DataSource = table;
                lister.DataSource = bSource;

            }
            else
            {

                string range = "";
                SqlDataAdapter DA = new SqlDataAdapter();
                string sqlSelectAll = invoice_query + " " + range + ";";
                DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                DataTable table = new DataTable();
                DA.Fill(table);

                BindingSource bSource = new BindingSource();
                bSource.DataSource = table;
                lister.DataSource = bSource;

            }

            con.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //create a new ExcelPackage
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //create 2 WorkSheets. One for the source data and one for the Pivot table
                ExcelWorksheet worksheetPivot = excelPackage.Workbook.Worksheets.Add("Pivot");
                ExcelWorksheet worksheetData = excelPackage.Workbook.Worksheets.Add("Data");

                //add some source data
                worksheetData.Cells["A1"].Value = "Column A";
                worksheetData.Cells["A2"].Value = "Group A";
                worksheetData.Cells["A3"].Value = "Group B";
                worksheetData.Cells["A4"].Value = "Group C";
                worksheetData.Cells["A5"].Value = "Group A";
                worksheetData.Cells["A6"].Value = "Group B";
                worksheetData.Cells["A7"].Value = "Group C";
                worksheetData.Cells["A8"].Value = "Group A";
                worksheetData.Cells["A9"].Value = "Group B";
                worksheetData.Cells["A10"].Value = "Group C";
                worksheetData.Cells["A11"].Value = "Group D";

                worksheetData.Cells["B1"].Value = "Column B";
                worksheetData.Cells["B2"].Value = "emc";
                worksheetData.Cells["B3"].Value = "fma";
                worksheetData.Cells["B4"].Value = "h2o";
                worksheetData.Cells["B5"].Value = "emc";
                worksheetData.Cells["B6"].Value = "fma";
                worksheetData.Cells["B7"].Value = "h2o";
                worksheetData.Cells["B8"].Value = "emc";
                worksheetData.Cells["B9"].Value = "fma";
                worksheetData.Cells["B10"].Value = "h2o";
                worksheetData.Cells["B11"].Value = "emc";

                worksheetData.Cells["C1"].Value = "Column C";
                worksheetData.Cells["C2"].Value = 299;
                worksheetData.Cells["C3"].Value = 792;
                worksheetData.Cells["C4"].Value = 458;
                worksheetData.Cells["C5"].Value = 299;
                worksheetData.Cells["C6"].Value = 792;
                worksheetData.Cells["C7"].Value = 458;
                worksheetData.Cells["C8"].Value = 299;
                worksheetData.Cells["C9"].Value = 792;
                worksheetData.Cells["C10"].Value = 458;
                worksheetData.Cells["C11"].Value = 299;

                worksheetData.Cells["D1"].Value = "Column D";
                worksheetData.Cells["D2"].Value = 40075;
                worksheetData.Cells["D3"].Value = 31415;
                worksheetData.Cells["D4"].Value = 384400;
                worksheetData.Cells["D5"].Value = 40075;
                worksheetData.Cells["D6"].Value = 31415;
                worksheetData.Cells["D7"].Value = 384400;
                worksheetData.Cells["D8"].Value = 40075;
                worksheetData.Cells["D9"].Value = 31415;
                worksheetData.Cells["D10"].Value = 384400;
                worksheetData.Cells["D11"].Value = 40075;

                //define the data range on the source sheet
                var dataRange = worksheetData.Cells[worksheetData.Dimension.Address];

                //create the pivot table
                var pivotTable = worksheetPivot.PivotTables.Add(worksheetPivot.Cells["B2"], dataRange, "PivotTable");

                //label field
                pivotTable.RowFields.Add(pivotTable.Fields["Column A"]);
                pivotTable.DataOnRows = false;

                //data fields
                var field = pivotTable.DataFields.Add(pivotTable.Fields["Column B"]);
                field.Name = "Count of Column B";
                field.Function = DataFieldFunctions.Count;

                field = pivotTable.DataFields.Add(pivotTable.Fields["Column C"]);
                field.Name = "Sum of Column C";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "0.00";

                field = pivotTable.DataFields.Add(pivotTable.Fields["Column D"]);
                field.Name = "Sum of Column D";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "€#,##0.00";

                FileInfo excelFile = new FileInfo(@"" + config.tempofiles + @"\PIVOT_TEST.xlsx");
                excelPackage.SaveAs(excelFile);
                System.Diagnostics.Process.Start(@"" + excelFile);

            }
        }
    }
}
