using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenezzInvoice
{
    public partial class arts : Form
    {
        public arts()
        {
            InitializeComponent();
        }
        //internal static string invoice_query = "SELECT  clave AS Clave, descr AS [Descripión], FORMAT( convert(numeric(18,5),replace(precio,',','')),'###,###,###.00000','ES-mx') AS [P.Venta], linea AS Linea, ume  AS [U.Medida], id FROM artsipl ORDER BY linea,clave ASC ";
        //internal static string invoice_querys = "SELECT  clave AS Clave, descr AS [Descripión], FORMAT(convert(numeric(18,5),replace(precio,',','')),'###,###,###.00000','ES-mx') AS [P.Venta], linea AS Linea, ume  AS [U.Medida], id FROM artsipl ";
        //internal static string invoice_query = "SELECT  clave AS Clave, descr AS [Descripión], precio AS [P.Venta], linea AS Linea, ume  AS [U.Medida], caja, mtscaja, kgspiece, kgscaja, costo AS Costo, id FROM artsipl ORDER BY clave ASC";

        internal static string invoice_query = "SELECT  clave AS Clave, descr AS [Descripión], precio AS [P.Venta], linea AS Linea, ume  AS [U.Medida], id FROM artsipl ORDER BY linea,clave ASC ";
        internal static string invoice_querys = "SELECT  clave AS Clave, descr AS [Descripión],precio AS [P.Venta], linea AS Linea, ume  AS [U.Medida], id FROM artsipl ";




        private void arts_Load(object sender, EventArgs e)
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

            lineas.Items.Add("");
            linee.Items.Add("");
            line.Items.Add("");
            lineadel.Items.Add("");
            string query = "SELECT nom FROM lineasipl ORDER BY nom ASC;";
            SqlCommand cm = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(cm);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    lineas.Items.Add("" + row["nom"]);
                    linee.Items.Add("" + row["nom"]);
                    line.Items.Add("" + row["nom"]);
                    lineadel.Items.Add("" + row["nom"]);
                }
            }
            da.Dispose(); cm.Dispose(); dt.Dispose();

            unide.Items.Add("");
            query = "SELECT ume FROM artsipl GROUP BY ume ORDER BY ume ASC;";
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    unide.Items.Add("" + row["ume"]);
                }
            }
            da.Dispose(); cm.Dispose(); dt.Dispose();

            unid.Items.Add("");
            query = "SELECT ume FROM artsipl GROUP BY ume ORDER BY ume ASC;";
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    unid.Items.Add("" + row["ume"]);
                }
            }
            da.Dispose(); cm.Dispose(); dt.Dispose();


            listap.Items.Add("");
            query = "SELECT nom FROM listasipl ORDER BY nom ASC;";
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    listap.Items.Add("" + row["nom"]);
                }
            }
            da.Dispose(); cm.Dispose(); dt.Dispose();

            resizegrid();
            con.Close();

            SqlCommand cmd = new SqlCommand("SELECT  clave FROM artsipl ORDER BY clave ASC;", con);
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
            lister.Columns[0].Width = 200;
            lister.Columns[1].Width = 300;
            lister.Columns[2].Width = 80;
            lister.Columns[3].Width = 100;
            lister.Columns[4].Width = 100;
            lister.Columns[5].Width = 50;
            lister.Columns[5].Visible = false;
        }

        private void arts_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void arts_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (idinvo.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea ELIMINAR?", "--Eliminar Producto --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    string cellValueid = idinvo.Text;
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "DELETE FROM artsipl WHERE id=" + cellValueid + ";";
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
                    cvee.Text = "";
                    descre.Text = "";
                    linee.Text = "";
                    unide.Text = "";
                    skue.Text = "";
                    precioe.Text = "";
                    costoe.Text = "";

                    cajae.Text = "";
                    m2cajae.Text = "";
                    kgpze.Text = "";
                    kgcajae.Text = "";
                    mede.Text = "";
                    pallete.Text = "";

                    SqlCommand cmd = new SqlCommand("SELECT  clave FROM artsipl ORDER BY clave ASC;", con);
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
            else { MessageBox.Show("Debes seleccionar un Producto para eliminarlo."); }

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            desclist.Text = "";
            pcolist.Text = "";
            listap.Text = "";
            SqlConnection con = new
            SqlConnection("" + config.cade);
            con.Open();

            if (cves.Text != "" || skus.Text != "" || lineas.Text != "")
            {




                string range = "";
                string cvess = "", skuss = "", lineass = "";
                if (cves.Text != "")
                {
                    cvess = " AND clave LIKE '%" + cves.Text + "%' ";
                }
                if (skus.Text != "")
                {
                    skuss = " AND sku  LIKE '%" + skus.Text + "%' ";
                }
                if (lineas.Text != "")
                {
                    lineass = " AND linea='" + lineas.Text + "' ";
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
                string sqlSelectAll = invoice_querys + " " + range + "  ORDER BY linea,clave ASC ;";
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

            cvee.Text = "";
            descre.Text = "";
            linee.Text = "";
            unide.Text = "";
            skue.Text = "";
            precioe.Text = "";
            costoe.Text = "";

            cajae.Text = "";
            m2cajae.Text = "";
            kgpze.Text = "";
            kgcajae.Text = "";
            mede.Text = "";
            pallete.Text = "";
            idinvo.Text = "";

        }

        private void lister_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            idinvo.Text = "";

            cvee.Text = "";
            descre.Text = "";
            linee.Text = "";
            unide.Text = "";
            skue.Text = "";
            precioe.Text = "";
            costoe.Text = "";

            cajae.Text = "";
            m2cajae.Text = "";
            kgpze.Text = "";
            kgcajae.Text = "";
            mede.Text = "";
            sizele.Text = "";
            desclist.Text = "";
            pcolist.Text = "";
            listap.Text = "";
            pallete.Text = "";


            if (e.RowIndex != -1)
            {

                var dataIndexNo = lister.Rows[e.RowIndex].Index.ToString();
                string cellValue = lister.Rows[e.RowIndex].Cells[0].Value.ToString();
                string cellValueid = lister.Rows[e.RowIndex].Cells[5].Value.ToString();

                idinvo.Text = "" + cellValueid;
                SqlConnection con = new
                SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT  * FROM artsipl WHERE id=" + cellValueid + ";";
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
                        linee.Text = "" + row["linea"];
                        unide.Text = "" + row["ume"];
                        skue.Text = "" + row["sku"];
                        //precioe.Text = "" + row["precio"];
                        //costoe.Text = "" + row["costo"];
                        cajae.Text = "" + row["caja"];
                        m2cajae.Text = "" + row["mtscaja"];
                        kgpze.Text = "" + row["kgspiece"];
                        kgcajae.Text = "" + row["kgscaja"];
                        mede.Text = "" + row["size"];
                        sizele.Text = "" + row["sizel"];
                        pallete.Text = "" + double.Parse("" + row["pallet"]);
                        precioe.Text = ""+ row["precio"];
                        costoe.Text = "" + row["costo"];




                        string queryexssi = "SELECT id, nom, obses FROM listasipl ORDER BY id ASC;";
                        SqlCommand cmpaasi = new SqlCommand(queryexssi, con);
                        SqlDataAdapter dapaasi = new SqlDataAdapter(cmpaasi);
                        DataTable dtpaasi = new DataTable();
                        dapaasi.Fill(dtpaasi);
                        int cuentapaasi = dtpaasi.Rows.Count;
                        if (cuentapaasi > 0)
                        {
                            foreach (DataRow rowp in dtpaasi.Rows)
                            {
                                string queryexssii = "SELECT id, clave, list FROM pricesipl WHERE clave = '" + row["clave"] + "' AND list = '" + rowp["id"] + "';";
                                SqlCommand cmpaasii = new SqlCommand(queryexssii, con);
                                SqlDataAdapter dapaasii = new SqlDataAdapter(cmpaasii);
                                DataTable dtpaasii = new DataTable();
                                dapaasii.Fill(dtpaasii);
                                int cuentapaasii = dtpaasii.Rows.Count;
                                if (cuentapaasii == 0)
                                {
                                    string qu = "INSERT INTO pricesipl (list, clave, precio) ";
                                    qu = qu + "VALUES('" + rowp["id"] + "','" + row["clave"] + "','" + row["precio"] + "');";
                                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                                    myCo.ExecuteNonQuery();
                                    myCo.Dispose();
                                }
                            }
                        }
                        cmpaasi.Dispose(); dapaasi.Dispose(); dtpaasi.Dispose();

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

               // precioe.Text = "" + double.Parse("" + precioe.Text).ToString("n2");
               // costoe.Text = "" + double.Parse("" + costoe.Text).ToString("n2");

                if (DialogResult.Yes == MessageBox.Show("¿Desea guardar?", " -- Almacenar cambios del producto --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "UPDATE artsipl  SET  descr='" + descre.Text + "', precio='" + precioe.Text + "', linea='" + linee.Text + "', ume='" + unide.Text + "', caja='" + cajae.Text + "', mtscaja='" + m2cajae.Text + "', kgspiece='" + kgpze.Text + "', kgscaja='" + kgcajae.Text + "', costo='" + costoe.Text + "', sku='" + skue.Text + "', size='" + mede.Text + "',sizel='" + sizele.Text + "', pallet='" + pallete.Text + "' ";
                    qu = qu + "WHERE id=" + idinvo.Text + ";";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();


                    //EVALUAR SI CREAR LINEA O NO
                    string query = "SELECT nom FROM lineasipl WHERE nom = '" + linee.Text + "';";
                    SqlCommand cm = new SqlCommand(query, con);
                    SqlDataAdapter da = new SqlDataAdapter(cm);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    int cuenta = 0;
                    cuenta = dt.Rows.Count;
                    if (cuenta == 0)
                    {
                        string queryi = "INSERT INTO lineasipl (nom) VALUES ('" + linee.Text + "');";
                        SqlCommand cmi = new SqlCommand(queryi, con);
                        cmi.ExecuteNonQuery();
                        cmi.Dispose();
                    }
                    da.Dispose(); cm.Dispose(); dt.Dispose();

                    lineas.Items.Clear();
                    linee.Items.Clear();
                    line.Items.Clear();
                    lineadel.Items.Clear();

                    lineas.Items.Add("");
                    linee.Items.Add("");
                    line.Items.Add("");
                    lineadel.Items.Add("");

                    query = "SELECT nom FROM lineasipl ORDER BY nom ASC;";
                    cm = new SqlCommand(query, con);
                    da = new SqlDataAdapter(cm);
                    dt = new DataTable();
                    da.Fill(dt);
                    cuenta = dt.Rows.Count;
                    if (cuenta > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            lineas.Items.Add("" + row["nom"]);
                            linee.Items.Add("" + row["nom"]);
                            line.Items.Add("" + row["nom"]);
                            lineadel.Items.Add("" + row["nom"]);
                        }
                    }
                    da.Dispose(); cm.Dispose(); dt.Dispose();


                    string range = " WHERE id =" + idinvo.Text + ";";
                    //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    //string sqlSelectAll = invoice_query + " " + range + "";
                    string sqlSelectAll = "SELECT clave AS Clave, descr AS [Descripión], precio AS [P.Venta], linea AS Linea, ume  AS [U.Medida], id FROM artsipl  " + range;
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = table;
                    lister.DataSource = bSource;
                    con.Close();

                    cvee.Text = "";
                    descre.Text = "";
                    linee.Text = "";
                    unide.Text = "";
                    skue.Text = "";
                    precioe.Text = "";
                    costoe.Text = "";

                    cajae.Text = "";
                    m2cajae.Text = "";
                    kgpze.Text = "";
                    kgcajae.Text = "";
                    mede.Text = "";
                    sizele.Text = "";
                    pallete.Text = "";
                }

            }
        }

        private void skue_KeyPress(object sender, KeyPressEventArgs e)
        {
            /*
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
            */
        }

        private void precioe_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void costoe_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void cajae_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void kgpze_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void m2cajae_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void kgcajae_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void sku_KeyPress(object sender, KeyPressEventArgs e)
        {
            /*
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
            */
        }

        private void precio_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void costo_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void caja_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void kgpz_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void m2caja_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void kgcaja_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (cve.Text != "")
            {
               // precio.Text = "" + double.Parse("" + precio.Text).ToString("n2");
               // costo.Text = "" + double.Parse("" + costo.Text).ToString("n2");


                if (DialogResult.Yes == MessageBox.Show("¿Desea guardar?", " -- Almacenar nuevo producto --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "INSERT INTO artsipl (clave, descr, precio, linea, ume, caja, mtscaja, kgspiece, kgscaja, costo, sku,size,sizel,pallet) ";
                    qu = qu + "VALUES('" + cve.Text + "','" + descr.Text + "','" + precio.Text + "','" + line.Text + "','" + unid.Text + "','" + caja.Text + "','" + m2caja.Text + "','" + kgpz.Text + "','" + kgcaja.Text + "','" + costo.Text + "','" + sku.Text + "','" + med.Text + "','" + sizel.Text + "','" + pallet.Text + "');SELECT SCOPE_IDENTITY();";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    string uuid = "" + myCo.ExecuteScalar().ToString();
                    myCo.Dispose();


                    //EVALUAR SI CREAR LINEA O NO
                    string query = "SELECT nom FROM lineasipl WHERE nom = '" + line.Text + "';";
                    SqlCommand cm = new SqlCommand(query, con);
                    SqlDataAdapter da = new SqlDataAdapter(cm);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    int cuenta = 0;
                    cuenta = dt.Rows.Count;
                    if (cuenta == 0)
                    {
                        string queryi = "INSERT INTO lineasipl (nom) VALUES ('" + line.Text + "');";
                        SqlCommand cmi = new SqlCommand(queryi, con);
                        cmi.ExecuteNonQuery();
                        cmi.Dispose();
                    }
                    da.Dispose(); cm.Dispose(); dt.Dispose();

                    lineas.Items.Clear();
                    linee.Items.Clear();
                    line.Items.Clear();
                    lineadel.Items.Clear();

                    lineas.Items.Add("");
                    linee.Items.Add("");
                    line.Items.Add("");
                    lineadel.Items.Add("");

                    query = "SELECT nom FROM lineasipl ORDER BY nom ASC;";
                    cm = new SqlCommand(query, con);
                    da = new SqlDataAdapter(cm);
                    dt = new DataTable();
                    da.Fill(dt);
                    cuenta = dt.Rows.Count;
                    if (cuenta > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            lineas.Items.Add("" + row["nom"]);
                            linee.Items.Add("" + row["nom"]);
                            line.Items.Add("" + row["nom"]);
                            lineadel.Items.Add("" + row["nom"]);
                        }
                    }
                    da.Dispose(); cm.Dispose(); dt.Dispose();




                    string range = " WHERE id =" + uuid + ";";
                    //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    string sqlSelectAll = "SELECT  clave AS Clave, descr AS [Descripión], precio AS [P.Venta], linea AS Linea, ume  AS [U.Medida], id FROM artsipl  " + range + "";
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = table;
                    lister.DataSource = bSource;


                    cve.Text = "";
                    descr.Text = "";
                    line.Text = "";
                    unid.Text = "";
                    sku.Text = "";
                    precio.Text = "";
                    costo.Text = "";

                    caja.Text = "";
                    m2caja.Text = "";
                    kgpz.Text = "";
                    kgcaja.Text = "";
                    med.Text = "";
                    sizel.Text = "";
                    pallet.Text = "";

                    cvee.Text = "";
                    descre.Text = "";
                    linee.Text = "";
                    unide.Text = "";
                    skue.Text = "";
                    precioe.Text = "";
                    costoe.Text = "";
                    sizele.Text = "";
                    cajae.Text = "";
                    m2cajae.Text = "";
                    kgpze.Text = "";
                    kgcajae.Text = "";
                    mede.Text = "";
                    pallete.Text = "";



                    SqlCommand cmd = new SqlCommand("SELECT clave FROM artsipl ORDER BY clave ASC;", con);
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
            else {
                MessageBox.Show("No puedes dejar la clave vacia");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (idinvo.Text == "")
            {
                MessageBox.Show("Debes seleccionar un articulo para poder cambiar sus valores en la lista de precios");
            }
            else
            {
                if (pcolist.Text != "")
                {
                    string lista = "";
                    string clave = "";


                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string queryexssi = "SELECT id FROM listasipl WHERE nom='" + listap.Text + "';";
                    SqlCommand cmpaasi = new SqlCommand(queryexssi, con);
                    SqlDataAdapter dapaasi = new SqlDataAdapter(cmpaasi);
                    DataTable dtpaasi = new DataTable();
                    dapaasi.Fill(dtpaasi);
                    int cuentapaasi = dtpaasi.Rows.Count;
                    if (cuentapaasi > 0)
                    {
                        foreach (DataRow rowp in dtpaasi.Rows)
                        {
                            lista = "" + rowp["id"];
                        }
                    }
                    cmpaasi.Dispose();
                    dtpaasi.Dispose();
                    cmpaasi.Dispose();

                    queryexssi = "SELECT clave FROM artsipl WHERE id=" + idinvo.Text + ";";
                    cmpaasi = new SqlCommand(queryexssi, con);
                    dapaasi = new SqlDataAdapter(cmpaasi);
                    dtpaasi = new DataTable();
                    dapaasi.Fill(dtpaasi);
                    cuentapaasi = dtpaasi.Rows.Count;
                    if (cuentapaasi > 0)
                    {
                        foreach (DataRow rowp in dtpaasi.Rows)
                        {
                            clave = "" + rowp["clave"];
                        }
                    }
                    cmpaasi.Dispose();
                    dtpaasi.Dispose();
                    cmpaasi.Dispose();


                    if (lista != "" && clave != "") {

                        //string qu = "UPDATE pricesipl SET precio='" + Math.Round(double.Parse("" + pcolist.Text), 2, MidpointRounding.AwayFromZero).ToString("n2") + "' WHERE clave='" + clave + "' AND list='" + lista + "'";
                        string qu = "UPDATE pricesipl SET precio='" +  pcolist.Text + "' WHERE clave='" + clave + "' AND list='" + lista + "'";
                        SqlCommand myCo = new SqlCommand(qu, config.conn);
                        myCo.ExecuteNonQuery();
                        myCo.Dispose();
                        if (desclist.Text != "")
                        {
                            qu = "UPDATE listasipl SET obses= '" + desclist.Text + "' WHERE id=" + lista + ";";
                            myCo = new SqlCommand(qu, config.conn);
                            myCo.ExecuteNonQuery();
                            myCo.Dispose();
                        }
                        //pcolist.Text = "" + Math.Round(double.Parse("" + pcolist.Text), 2, MidpointRounding.AwayFromZero).ToString("n2");
                    }
                    con.Close();
                }

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (nvalistn.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea guardar una NUEVA LISTA?\r\nPrecione ACEPTAR para crear el listado.", "--Almacenar nueva lista de precios--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {


                    string idlist = "", cve = "";
                    SqlConnection con = new SqlConnection(config.cade);
                    string queryexss = "SELECT id,nom,obses FROM listasipl WHERE nom='" + nvalistn.Text + "';";
                    SqlCommand cmpaas = new SqlCommand(queryexss, con);
                    SqlDataAdapter dapaas = new SqlDataAdapter(cmpaas);
                    DataTable dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    int cuentapaas = dtpaas.Rows.Count;
                    if (cuentapaas > 0)
                    {
                        MessageBox.Show("Esta nombre de lista ya esta en uso.");
                    }
                    else
                    {
                        string uuid = "";
                        string qu = "INSERT INTO listasipl (nom, obses) ";
                        qu = qu + "VALUES('" + nvalistn.Text + "','" + descrpnval.Text + "');SELECT SCOPE_IDENTITY();";
                        SqlCommand myCo = new SqlCommand(qu, config.conn);
                        uuid = "" + myCo.ExecuteScalar().ToString();
                        myCo.Dispose();

                        string queryexssi = "SELECT clave, descr, precio FROM artsipl ORDER BY clave ASC;";
                        SqlCommand cmpaasi = new SqlCommand(queryexssi, con);
                        SqlDataAdapter dapaasi = new SqlDataAdapter(cmpaasi);
                        DataTable dtpaasi = new DataTable();
                        dapaasi.Fill(dtpaasi);
                        int cuentapaasi = dtpaasi.Rows.Count;
                        if (cuentapaasi > 0)
                        {
                            foreach (DataRow rowp in dtpaasi.Rows)
                            {
                                qu = "INSERT INTO pricesipl (list, clave, precio) ";
                                qu = qu + "VALUES('" + uuid + "','" + rowp["clave"] + "','" + rowp["precio"] + "');SELECT SCOPE_IDENTITY();";
                                myCo = new SqlCommand(qu, config.conn);
                                myCo.ExecuteNonQuery();
                                myCo.Dispose();
                            }
                        }


                        listap.Items.Clear();
                        listap.Items.Add("");
                        string query = "SELECT nom FROM listasipl ORDER BY nom ASC;";
                        SqlCommand cm = new SqlCommand(query, con);
                        SqlDataAdapter da = new SqlDataAdapter(cm);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        int cuenta = dt.Rows.Count;
                        if (cuenta > 0)
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                listap.Items.Add("" + row["nom"]);
                            }
                        }
                        da.Dispose(); cm.Dispose(); dt.Dispose();


                        nvalistn.Text = ""; descrpnval.Text = "";
                    }
                    cmpaas.Dispose(); dapaas.Dispose(); dtpaas.Dispose();

                }
            }
        }

        private void pcolist_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void listap_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (idinvo.Text != "")
            {
                string idlist = "", cve = "";

                if (listap.Text != "")
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    string queryexss = "SELECT id,nom,obses FROM listasipl WHERE nom='" + listap.Text + "';";
                    SqlCommand cmpaas = new SqlCommand(queryexss, con);
                    SqlDataAdapter dapaas = new SqlDataAdapter(cmpaas);
                    DataTable dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    int cuentapaas = dtpaas.Rows.Count;
                    if (cuentapaas > 0)
                    {
                        foreach (DataRow rowp in dtpaas.Rows)
                        {
                            idlist = "" + rowp["id"].ToString();
                            desclist.Text = "" + rowp["obses"].ToString();
                        }
                    }
                    cmpaas.Dispose(); dapaas.Dispose(); dtpaas.Dispose();


                    con = new SqlConnection(config.cade);
                    queryexss = "SELECT clave FROM artsipl WHERE id=" + idinvo.Text + ";";
                    cmpaas = new SqlCommand(queryexss, con);
                    dapaas = new SqlDataAdapter(cmpaas);
                    dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    cuentapaas = dtpaas.Rows.Count;
                    if (cuentapaas > 0)
                    {
                        foreach (DataRow rowp in dtpaas.Rows)
                        {
                            //listpr.Items.Add(rowp["nom"].ToString());
                            cve = "" + rowp["clave"].ToString();
                        }
                    }
                    cmpaas.Dispose(); dapaas.Dispose(); dtpaas.Dispose();


                    con = new SqlConnection(config.cade);
                    queryexss = "SELECT  id, list, clave, precio  FROM pricesipl WHERE list='" + idlist + "' AND clave='" + cve + "';";
                    cmpaas = new SqlCommand(queryexss, con);
                    dapaas = new SqlDataAdapter(cmpaas);
                    dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    cuentapaas = dtpaas.Rows.Count;
                    double montoss = 0;
                    if (cuentapaas > 0)
                    {
                        foreach (DataRow rowp in dtpaas.Rows)
                        {
                            //pcolist.Text = "" + rowp["precio"];
                            //montoss = "" 
                            pcolist.Text = "" + rowp["precio"].ToString();
                        }
                    }
                    cmpaas.Dispose(); dapaas.Dispose(); dtpaas.Dispose();
                    con.Close();
                }
                else
                {
                    pcolist.Text = "";
                    desclist.Text = "";
                }


            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("¿Desea eliminar esta LISTA?\r\nPrecione ACEPTAR para eliminar el listado.", "--Almacenar eliminar lista de precios--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                if (listap.Text != "")
                {
                    string idlist = "";
                    SqlConnection con = new SqlConnection(config.cade);
                    string queryexss = "SELECT id,nom,obses FROM listasipl WHERE nom='" + listap.Text + "';";
                    SqlCommand cmpaas = new SqlCommand(queryexss, con);
                    SqlDataAdapter dapaas = new SqlDataAdapter(cmpaas);
                    DataTable dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    int cuentapaas = dtpaas.Rows.Count;
                    if (cuentapaas > 0)
                    {
                        foreach (DataRow rowp in dtpaas.Rows)
                        {
                            idlist = "" + rowp["id"].ToString();
                        }
                    }
                    cmpaas.Dispose(); dapaas.Dispose(); dtpaas.Dispose();

                    string qu = "DELETE FROM pricesipl WHERE list='" + idlist + "';";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();
                    con.Close();

                    qu = "DELETE FROM listasipl WHERE id=" + idlist + ";";
                    myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    listap.Items.Clear();
                    listap.Items.Add("");
                    string query = "SELECT nom FROM listasipl ORDER BY nom ASC;";
                    SqlCommand cm = new SqlCommand(query, con);
                    SqlDataAdapter da = new SqlDataAdapter(cm);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    int cuenta = dt.Rows.Count;
                    if (cuenta > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            listap.Items.Add("" + row["nom"]);
                        }
                    }
                    da.Dispose(); cm.Dispose(); dt.Dispose();


                }

            }
        }

        private void pallete_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void pallet_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void mede_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void sizele_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void med_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void sizel_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {

            if ( lineadel.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea eliminar linea?", " -- Eliminar linea del catálogo --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    //EVALUAR SI CREAR LINEA O NO
                    string query = "SELECT linea FROM artsipl WHERE linea = '" + lineadel.Text + "';";
                    SqlCommand cm = new SqlCommand(query, con);
                    SqlDataAdapter da = new SqlDataAdapter(cm);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    int cuenta = 0;
                    cuenta = dt.Rows.Count;

                            if (cuenta == 0)
                            {
                                string queryi = "DELETE FROM lineasipl WHERE nom='" + lineadel.Text + "';";
                                SqlCommand cmi = new SqlCommand(queryi, con);
                                cmi.ExecuteNonQuery();
                                cmi.Dispose();
                                lineas.Items.Clear();
                                linee.Items.Clear();
                                line.Items.Clear();
                                lineadel.Items.Clear();

                                lineas.Items.Add("");
                                linee.Items.Add("");
                                line.Items.Add("");
                                lineadel.Items.Add("");

                                query = "SELECT nom FROM lineasipl ORDER BY nom ASC;";
                                cm = new SqlCommand(query, con);
                                da = new SqlDataAdapter(cm);
                                dt = new DataTable();
                                da.Fill(dt);
                                cuenta = dt.Rows.Count;
                                if (cuenta > 0)
                                {
                                    foreach (DataRow row in dt.Rows)
                                    {
                                        lineas.Items.Add("" + row["nom"]);
                                        linee.Items.Add("" + row["nom"]);
                                        line.Items.Add("" + row["nom"]);
                                        lineadel.Items.Add("" + row["nom"]);
                                    }
                                }
                            }
                            else
                    {
                        MessageBox.Show("La linea aun tiene articulos dentro de ella.\r\nCámbielos o elimine los productos", "Eliminación de lineas");
                    }
                            da.Dispose(); cm.Dispose(); dt.Dispose();
                            con.Close();
              }
           } else
            {
                MessageBox.Show("Debe seleccionar una linea para poder eliminarla", "Eliminación de lineas");
            }
        }

        private void lineadel_SelectedIndexChanged(object sender, EventArgs e)
        {
            lincuan.Text = "";
            if (lineadel.Text != "")
            {
                SqlConnection con = new SqlConnection(config.cade);
                con.Open();
                //EVALUAR SI CREAR LINEA O NO
                string query = "SELECT linea FROM artsipl WHERE linea = '" + lineadel.Text + "';";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = 0;
                cuenta = dt.Rows.Count;
                lincuan.Text = "Articulos en la linea: " + cuenta;
                con.Close();
            }
         }
    }
}
