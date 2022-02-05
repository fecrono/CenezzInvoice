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
    public partial class add : Form
    {
        public int deci = 0;
        public add()
        {
            InitializeComponent();
        }

        private void add_Load(object sender, EventArgs e)
        {
                SqlConnection con = new SqlConnection(config.cade);
                SqlCommand cmd = new SqlCommand("SELECT  clave FROM artsipl ORDER BY clave ASC;", con);
                con.Open();

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



            query = "SELECT TOP(1)* FROM configs ORDER BY id ASC;";
            SqlCommand cm = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(cm);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    obs1.Text = "" + row["obs1"];
                    obs2.Text = "" + row["obs2"];
                    obs3.Text = "" + row["obs3"];
                    obs4.Text = "" + row["obs4"];
                    obs5.Text = "" + row["obs5"];
                }
            }
            dt.Dispose(); cm.Dispose(); da.Dispose();



            SqlDataReader reader = cmd.ExecuteReader();
                AutoCompleteStringCollection MyCollection = new AutoCompleteStringCollection();
                while (reader.Read())
                {
                    MyCollection.Add(reader.GetString(0));
                }
                cveadd.AutoCompleteCustomSource = MyCollection;
                    cmd.Dispose(); reader.Dispose();


            listpr.Items.Add("");
            string queryexss = "SELECT nom FROM listasipl ORDER BY nom ASC;";
            SqlCommand cmpaas = new SqlCommand(queryexss, con);
            SqlDataAdapter dapaas = new SqlDataAdapter(cmpaas);
            DataTable dtpaas = new DataTable();
            dapaas.Fill(dtpaas);
            int cuentapaas = dtpaas.Rows.Count;
            if (cuentapaas > 0)
            {
                foreach (DataRow rowp in dtpaas.Rows)
                {
                    listpr.Items.Add(rowp["nom"].ToString());
                }
            }
            cmpaas.Dispose(); dapaas.Dispose(); dtpaas.Dispose();

            cmd = new SqlCommand("SELECT clave FROM gastos ORDER BY clave ASC;", con);
            reader = cmd.ExecuteReader();
            AutoCompleteStringCollection MyCollectionS = new AutoCompleteStringCollection();
            while (reader.Read())
            {
                MyCollectionS.Add(reader.GetString(0));
            }
            claveserv.AutoCompleteCustomSource = MyCollectionS;
            cmd.Dispose(); reader.Dispose();

            deci = Int32.Parse("" + deces.Value);
            con.Close();


        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
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

        private void button8_Click(object sender, EventArgs e)
        {
            containeradd.Items.Clear();
            string toadd = "";
            string addded = "yes";
            string itac = "";
            toadd = "" + containsadd.Text;
            if (toadd != "")
            {
                if (addded == "yes")
                {

                    contains.Items.Add("" + toadd);
                    containsadd.Text = "";
                    foreach (object item in contains.Items)
                    {
                        itac = "" + item.ToString();
                        if (itac == "" + toadd)
                        {
                            addded = "no";
                        }
                        containeradd.Items.Add("" + item.ToString());
                    }
                }
                else { MessageBox.Show("Este contenedor ya esta en la lista."); }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

            if (contains.SelectedIndex != -1)
            {
                int countitems = 0;
                string selectedItem = contains.Items[contains.SelectedIndex].ToString();
                contains.Items.RemoveAt(contains.SelectedIndex);
                containeradd.Items.Clear();

                foreach (ListViewItem item in addrows.Items)
                {
                    if (selectedItem == "" + item.SubItems[5].Text)
                    {
                        addrows.Items[countitems].SubItems[5].Text = "";
                    }
                    else
                    {
                        /*
                        queryinsref = "UPDATE containersipl SET container ='" + updtc.Text + "',precinto='" + anteriorrecinto + "' WHERE ord ='" + config.idinvoice + "' AND container ='" + anterior + "';";
                        myCommandref = new SqlCommand(queryinsref, config.conn);
                        myCommandref.ExecuteNonQuery();
                        myCommandref.Dispose();
                        */
                    }
                    countitems = countitems + 1;
                }


                foreach (object item in contains.Items)
                {
                    containeradd.Items.Add("" + item.ToString());
                }
            }
            else
            {
                MessageBox.Show("Seleccione un contenedor de la lista para quitarlo.");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (emit.Text != "")
            {
                SqlConnection con = new SqlConnection("" + config.cade);
                SqlCommand cmd = new SqlCommand("SELECT  clave FROM artsipl ORDER BY clave ASC;", con);
                con.Open();
                    string query = "SELECT nom FROM empresasipl WHERE id =" + emit.Text + ";";
                    SqlCommand cm = new SqlCommand(query, con);
                    SqlDataAdapter da = new SqlDataAdapter(cm);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    int cuenta = dt.Rows.Count;
                    if (cuenta > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {

                        /*
                        * radio.Text = Convet.ToString(row["NombreEntorno"]);
                        * }
                        */
                            nomemp.Text = "" + row["nom"];
                        }
                    }
                    else
                    {
                        emit.Text = "";
                        MessageBox.Show("Empresa no existe");
                    }
                    con.Close();
                
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (recep.Text != "")
            {
                recnom.Text = "";
                dirrec.Text = "";
                nume.Text = "";
                numi.Text = "";
                col.Text = "";
                cd.Text = "";
                mun.Text = "";
                edo.Text = "";
                pais.Text = "";
                niff.Text = "";

                SqlConnection con = new SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT c.id, c.nom, c.nif, c.calle, c.num, c.numi, c.col, c.mun, c.cd, c.edo, c.pais, cp,precios,l.nom as nompre  FROM clientesipl AS c  INNER JOIN    listasipl as l ON l.id=c.precios  WHERE c.id =" + recep.Text + ";";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        recnom.Text = "" + row["nom"];
                        dirrec.Text = "" + row["calle"];
                        nume.Text = "" + row["num"];
                        numi.Text = "" + row["numi"];
                        col.Text = "" + row["col"];
                        cd.Text = "" + row["cd"];
                        mun.Text = "" + row["mun"];
                        edo.Text = "" + row["edo"];
                        pais.Text = "" + row["pais"];
                        niff.Text = "" + row["nif"];

                        listpr.Text = "" + row["nompre"];
                    }
                }
                else
                {
                    recep.Text = "";
                    MessageBox.Show("Empresa no existe");
                }
                con.Close();

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if ((cveadd.Text != "") && (cantadd.Text != "") && (containeradd.Text !=""))
            {


                string adesc, clave, price, umes = "", contain,pallet="";
                clave = "" + cveadd.Text;
                double acant1 = 0, aprecio1 = 0, kgspiece=0, totalon = 0, pallets = 0, kgscaja = 0, kilos = 0,metros=0, cajas=0;


                acant1 = Math.Round(double.Parse(cantadd.Text), 2, MidpointRounding.AwayFromZero);

                SqlConnection con = new SqlConnection(config.cade);
                SqlCommand cmd = new SqlCommand("SELECT clave,descr,precio,linea,ume,pallet,kgspiece,kgscaja,mtscaja FROM artsipl WHERE clave= '" + clave + "' ORDER BY clave ASC;", con);
                con.Open();
                SqlDataAdapter dapaas = new SqlDataAdapter(cmd);
                DataTable dtpaas = new DataTable();
                dapaas.Fill(dtpaas);
                int cuentapaas = dtpaas.Rows.Count;
                if (cuentapaas > 0)
                {
                    foreach (DataRow rowp in dtpaas.Rows)
                    {
                        aprecio1 = Math.Round(double.Parse("" + rowp["precio"]), deci, MidpointRounding.AwayFromZero);
                        umes = "" + rowp["ume"];
                        pallet = "" + rowp["pallet"];
                        kgspiece = Math.Round(double.Parse("" + rowp["kgspiece"]), 2, MidpointRounding.AwayFromZero);
                        kgscaja = Math.Round(double.Parse("" + rowp["kgscaja"]), 2, MidpointRounding.AwayFromZero);
                        metros = Math.Round(double.Parse("" + rowp["mtscaja"]), 2, MidpointRounding.AwayFromZero);
                        //kilos = Math.Round(double.Parse("" + rowp["precio"]), 2, MidpointRounding.AwayFromZero);
                    }
                }

                dapaas.Dispose(); dtpaas.Dispose(); cmd.Dispose();

                if (listpr.Text != "")
                {
                    cmd = new SqlCommand("SELECT l.id,l.nom,p.precio FROM listasipl as l INNER JOIN pricesipl as p ON p.list = l.id WHERE p.clave='" + clave + "' AND l.nom='"+ listpr.Text + "';", con);
                    //MessageBox.Show("" + "SELECT l.id,l.nom,p.precio FROM listasipl as l INNER JOIN pricesipl as p ON p.list = l.id WHERE p.clave='" + clave + "' AND l.nom='" + listpr.Text + "';");
                    dapaas = new SqlDataAdapter(cmd);
                    dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    cuentapaas = dtpaas.Rows.Count;
                    if (cuentapaas > 0)
                    {
                        foreach (DataRow rowp in dtpaas.Rows)
                        {
                            aprecio1 = Math.Round(double.Parse("" + rowp["precio"]),deci, MidpointRounding.AwayFromZero);
                        }
                    }
                    dapaas.Dispose(); dtpaas.Dispose(); cmd.Dispose();
                }
                con.Close();


                try { pallets = double.Parse("" + pallet); }
                catch { pallets = 0; }

                try
                {
                    pallets = double.Parse("" + acant1) / pallets;
                }
                catch { pallets = 1; }
                //if (pallets < 1) { pallets = 1; }
                if (pallets.ToString() == "∞") { pallets = 0; }

                try
                {
                    cajas = double.Parse("" + acant1) / metros;
                }
                catch { cajas = 1; }
                //if (pallets < 1) { pallets = 1; }
                if (cajas.ToString() == "∞") { cajas = 0; }

                try
                {
                    kilos = double.Parse("" + cajas) * kgscaja;
                    if (kilos == 0)
                    {
                        kilos = double.Parse("" + acant1) * kgspiece;
                    }
                    if (kilos == 0)
                    {
                        kilos = 0;
                    }
                }
                catch { kilos = 1; }
                //if (pallets < 1) { pallets = 1; }
                if (kilos.ToString() == "∞") { kilos = 0; }



                totalon = aprecio1 * acant1;
                //cant clave       ume         punit          kilos caja        importe                           container
                string[] row1 = { "" + clave, "" + umes, kgscaja.ToString("n2"), "" + aprecio1.ToString("n"  +deci), "" + totalon.ToString("n" + +deci), "" + containeradd.Text, "" + kilos.ToString("n2"), "" +  pallets.ToString("n2") };
                addrows.Items.Add("" + acant1.ToString("n2")).SubItems.AddRange(row1);


                //Sumar los elementos del listbox
                Double dblSuma = 0;

                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[5].Text);
                }
                totrefs.Text = "" + dblSuma.ToString("n" + +deci);

                dblSuma = 0;
                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[7].Text);
                }
                neto.Text = "" + dblSuma.ToString("n"  + deci );
                dblSuma = 0;

                //obtener peso bruto
                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[8].Text)* Convert.ToDouble(palletkgs.Text);
                }
                dblSuma = dblSuma + Convert.ToDouble(neto.Text);
                bruto.Text = "" + dblSuma.ToString("n2");
                dblSuma = 0;


                //calcular los totales en peso de cada cont
                logcontainer.Text = "";
                foreach (object item in contains.Items)
                {
                    dblSuma = 0;
                    foreach (ListViewItem items in addrows.Items)
                    {
                        if (item.ToString() + "" == "" + items.SubItems[6].Text)
                        {
                            dblSuma += Convert.ToDouble(items.SubItems[7].Text);
                        }
                    }
                    if (dblSuma > 0)
                    {
                        logcontainer.AppendText("-" + item.ToString() + ": " + dblSuma.ToString("n2") + " kgs.\r\n");
                    }
                    else {
                        logcontainer.AppendText("-" + item.ToString() + ": " + dblSuma.ToString("n2") + " kgs.\r\n");
                    }
                }


                cantadd.Text = "1.0";
                //Limpar Arts de los textboxs
                //acant.Text = "1"; adesc.Text = ""; apart.Text = ""; aprecio.Text = ""; sistrefadd.Text = ""; sistmanadd.Text = ""; cvesatpa.Text = "25101500";
            } else
            {
                MessageBox.Show("No se puede agregar este concepto sin llenar todos los campos");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int sele = 0;
            double totalizar, subbo;
            sele = sele + int.Parse(idadded.Text);
            Double dblSuma = 0;
            if (idadded.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea eliminar?", "--Eliminar partida--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    this.addrows.Items[sele].Remove();
                    idadded.Text = "";
                    //Sumar los elementos del listbox
                    foreach (ListViewItem item in addrows.Items)
                    {
                        dblSuma += Convert.ToDouble(item.SubItems[5].Text);
                    }


                    foreach (ListViewItem item in serves.Items)
                    {
                        dblSuma += Convert.ToDouble(item.SubItems[4].Text);
                    }
                    totrefs.Text = "" + dblSuma.ToString("n"  +deci);

                    dblSuma = 0;
                    foreach (ListViewItem item in addrows.Items)
                    {
                        dblSuma += Convert.ToDouble(item.SubItems[7].Text);
                    }
                    neto.Text = "" + dblSuma.ToString("n2");
                    dblSuma = 0;

                    //obtener peso bruto
                    foreach (ListViewItem item in addrows.Items)
                    {
                        dblSuma += Convert.ToDouble(item.SubItems[8].Text) * Convert.ToDouble(palletkgs.Text);
                    }
                    dblSuma = dblSuma + Convert.ToDouble(neto.Text);
                    bruto.Text = "" + dblSuma.ToString("n2");
                    dblSuma = 0;

                    //calcular los totales en peso de cada cont
                    logcontainer.Text = "";
                    foreach (object item in contains.Items)
                    {
                        dblSuma = 0;
                        foreach (ListViewItem items in addrows.Items)
                        {
                            if (item.ToString() + "" == "" + items.SubItems[6].Text)
                            {
                                dblSuma += Convert.ToDouble(items.SubItems[7].Text);
                            }
                        }
                        if (dblSuma > 0)
                        {
                            logcontainer.AppendText("-" + item.ToString() + ": " + dblSuma.ToString("n2") + " kgs.\r\n");
                        }
                        else
                        {
                            logcontainer.AppendText("-" + item.ToString() + ": 0 kgs.\r\n");
                        }
                    }

                    /////////////////Eliminarlo
                    delref.Enabled = false;
                }
            }
            idadded.Text = "";
        }

        private void addrows_SelectedIndexChanged(object sender, EventArgs e)
        {
            cantedit.Text = "";
            clavedit.Text = "";
            idadded.Text = "";
            palled.Text = "";
            kgboxx.Text = "";
            kilose.Text = "";
            if (addrows.SelectedIndices.Count <= 0)
            {
                return;
            }
            int intselectedindex = addrows.SelectedIndices[0];
            if (intselectedindex >= 0)
            {
                idadded.Text = "" + intselectedindex;
                //extrae el contenido del campo: MessageBox.Show(addreflist.Items[intselectedindex].Text);
                cantedit.Text = "" + addrows.Items[intselectedindex].Text;
                clavedit.Text = "" + addrows.Items[intselectedindex].SubItems[1].Text;
                cue.Text = "" + addrows.Items[intselectedindex].SubItems[4].Text;
                kgboxx.Text = "" + addrows.Items[intselectedindex].SubItems[3].Text;
                kilose.Text = "" + addrows.Items[intselectedindex].SubItems[7].Text;
                containeradd.Text = "" + addrows.Items[intselectedindex].SubItems[6].Text;
                palled.Text = "" + addrows.Items[intselectedindex].SubItems[8].Text;
                delref.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void emit_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void add_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void add_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            double dblSuma = 0,brutocont=0;
            if (currency.Text == "")
            {
                MessageBox.Show("No se puede guardar el Invoice sin seleccionar una MONEDA.");
            }
            else
            {
                if (folio.Text == "")
                {
                    MessageBox.Show("No se puede guardar el Invoice sin Colocar un folio para este.");
                }
                else
                {
                    if (DialogResult.Yes == MessageBox.Show("¿Desea guardar?\r\nInvoice/packing list #: " + folio.Text + ".", "--Almacenar Set de Documentos--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                    {

                        DateTime Hoy = DateTime.Today;
                        DateTime dte = Hoy;
                        string foliado = "" + dte.ToString("yyyy");
                        string query = "SELECT year,folio FROM folios WHERE year ='" + foliado + "';";
                        SqlCommand cm = new SqlCommand(query, config.conn);
                        SqlDataAdapter da = new SqlDataAdapter(cm);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        int cuenta = dt.Rows.Count;
                        int foloact = 0;
                        if (cuenta > 0)
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                foloact = int.Parse("" + row["folio"].ToString());
                                foloact = foloact + 1;
                                foliado = "" + foloact.ToString() + "-" + foliado;

                                string quin = "UPDATE folios SET folio='" + foloact.ToString() + "' ";
                                quin = quin + "WHERE year = '" + dte.ToString("yyyy") + "';";
                                SqlCommand myCoi = new SqlCommand(quin, config.conn);
                                myCoi.ExecuteNonQuery();
                                myCoi.Dispose();
                            }
                        }
                        else
                        {
                            string quin = "INSERT INTO  folios (year, folio) ";
                            quin = quin + "VALUES('" + dte.ToString("yyyy") + "','1');";
                            SqlCommand myCoi = new SqlCommand(quin, config.conn);
                            myCoi.ExecuteNonQuery();
                            myCoi.Dispose();
                            foliado = "1-" + dte.ToString("yyyy");
                        }
                        cm.Dispose(); da.Dispose(); dt.Dispose();

                        string uuid = "";
                        SqlConnection con = new SqlConnection(config.cade);
                        con.Open();
                        string qu = "INSERT INTO invoicespl (folio,number, empresa, nomcli, callecli, numcli, numclii, colcli, muncli, edocli, paiscli, fecha, albaran, origdest,tot,idcli,currency,obs1,obs2,obs3,obs4,obs5,pesoneto,pesobruto) ";
                        qu = qu + "VALUES('" + folio.Text + "','" + foliado + "','" + emit.Text + "','" + recnom.Text + "','" + dirrec.Text + "','" + nume.Text + "','" + numi.Text + "','" + col.Text + "','" + mun.Text + "','" + edo.Text + "','" + pais.Text + "','" + fecha.Text + "','" + alba.Text + "','" + oridest.Text + "','" + totrefs.Text + "','" + recep.Text + "','" + currency.Text + "','" + obs1.Text + "','" + obs2.Text + "','" + obs3.Text + "','" + obs4.Text + "','" + obs5.Text + "','" + neto.Text + "','" + bruto.Text + "');SELECT SCOPE_IDENTITY();";
                        SqlCommand myCo = new SqlCommand(qu, config.conn);
                        uuid = "" + myCo.ExecuteScalar().ToString();
                        myCo.Dispose();

                        string itac = "";
                        foreach (object item in contains.Items)
                        {
                            itac = "" + item.ToString();
                           
                            if (itac != "")
                            {
                                qu = "INSERT INTO containersipl (ord,container) VALUES ('" + uuid + "','" + itac + "');";
                                myCo = new SqlCommand(qu, config.conn);
                                myCo.ExecuteNonQuery();
                                myCo.Dispose();

// aqui meter el update para el total de kilos del containerrecien insertado
                                dblSuma = 0; brutocont = 0;
                                foreach (ListViewItem items in addrows.Items)
                                {
                                    if (itac + "" == "" + items.SubItems[6].Text)
                                    {
                                        dblSuma += Convert.ToDouble(items.SubItems[7].Text);
                                        brutocont += Convert.ToDouble(items.SubItems[8].Text);
                                    }
                                }
                                if (dblSuma > 0)
                                {
                                    brutocont = (Convert.ToDouble("" + palletkgs.Text) * brutocont) + dblSuma;
                                    qu = "UPDATE containersipl SET  pesoneto='"+ dblSuma.ToString("n2") + "', pesobruto='"+ brutocont.ToString("n2") + "' WHERE container= '" + itac + "';";
                                    myCo = new SqlCommand(qu, config.conn);
                                    myCo.ExecuteNonQuery();
                                    myCo.Dispose();
                                }
//fin de kilos en container

                            }
                        }

                        string queryinsref = ""; brutocont = 0;
                        foreach (ListViewItem item in addrows.Items)
                        {                                    // 0    1    2      4      5      6         7     8
                                                             //cant,clave,ume, punit,importe,container,kgstot,pallets
                            brutocont = (Convert.ToDouble("" + palletkgs.Text) * Convert.ToDouble("" + item.SubItems[8].Text)) + Convert.ToDouble(item.SubItems[7].Text);
                            queryinsref = "INSERT INTO rowsipl ( ord, cant, clave, ume, pu, importe, container,pallets,pesoneto, pesobruto)" +
                            " VALUES ('" + uuid + "', '" + item.SubItems[0].Text + "', '" + item.SubItems[1].Text + "', '" + item.SubItems[2].Text + "', '" + item.SubItems[4].Text + "', '" + item.SubItems[5].Text + "', '" + item.SubItems[6].Text + "','" + item.SubItems[8].Text + "','" + item.SubItems[7].Text + "','" + brutocont.ToString("n2") + "');";
                                                            // 0                                      1                                      2                              4                            5                                      6                        7                                  8
                                                            //cant,                            clave,                                 ume,                         punit,                        importe,                              container,                  kgstot,                          pallets
                                SqlCommand myCommandref = new SqlCommand(queryinsref, con);
                                myCommandref.ExecuteNonQuery();
                                myCommandref.Dispose();
                            }

                            queryinsref = "";
                            foreach (ListViewItem items in serves.Items)
                            {
                                queryinsref = "INSERT INTO rowsservpl (  ord, cant, clave, descrip, cu,  total) VALUES ('" + uuid + "', '" + items.SubItems[0].Text + "', '" + items.SubItems[1].Text + "', '" + items.SubItems[2].Text + "', '" + items.SubItems[3].Text + "', '" + items.SubItems[4].Text + "');";
                                SqlCommand myCommandref = new SqlCommand(queryinsref, con);
                                myCommandref.ExecuteNonQuery();
                                myCommandref.Dispose();
                            }

                            con.Close();
                            this.Close();

                        }

                    }
                }
            }
        

        private void button2_Click(object sender, EventArgs e)
        {
            string conten = "";
            Random rnd = new Random();
            int rcant = 0;
            rcant = int.Parse("" + randcant.Text);
            containeradd.Items.Clear();
            contains.Items.Clear();
            for (int inn = 1; inn <= rcant; inn++)
            {
                
                for (int i = 1; i <= 4; i++)
                {
                    char randomChar = (char)rnd.Next('a', 'z');
                    conten = conten + "" + randomChar;
                }
                conten = conten.ToUpper();
                for (int i = 1; i <= 7; i++)
                {
                    int dice = rnd.Next(1, 10);
                    conten = conten + "" + dice;
                }
                //MessageBox.Show( inn  +" - " + conten);
                containeradd.Items.Add("" + conten);
                contains.Items.Add("" + conten);
                conten = "";
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

        private void addserv_Click(object sender, EventArgs e)
        {
            string sclave;
            string descripte = "";
            if ((claveserv.Text != "") && (cantserv.Text != "") && (precserv.Text != ""))
            {


                sclave = "" + claveserv.Text;
                double cants = 0, precios = 0, totalon = 0;
                cants = Math.Round(double.Parse(cantserv.Text), 2, MidpointRounding.AwayFromZero);
                precios = Math.Round(double.Parse(precserv.Text), deci, MidpointRounding.AwayFromZero);

                SqlConnection con = new SqlConnection(config.cade);
                SqlCommand cmd = new SqlCommand("SELECT id,clave,descr,precio FROM gastos WHERE clave= '" + sclave + "' ORDER BY clave ASC;", con);
                con.Open();
                SqlDataAdapter dapaas = new SqlDataAdapter(cmd);
                DataTable dtpaas = new DataTable();
                dapaas.Fill(dtpaas);
                int cuentapaas = dtpaas.Rows.Count;
                if (cuentapaas > 0)
                {
                    foreach (DataRow rowp in dtpaas.Rows)
                    {
                        descripte = "" + rowp["descr"];
                    }
                }

                con.Close();

                totalon = precios * cants;
                //cant clave       ume         punit                         importe                           container
                string[] row1 = { "" + sclave, "" + descripte, "" + precios.ToString("n" + deci), "" + totalon.ToString("n" +  deci) };
                serves.Items.Add("" + cants.ToString("n2")).SubItems.AddRange(row1);


                //Sumar los elementos del listbox
                Double dblSuma = 0;

                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[5].Text);
                }


                foreach (ListViewItem item in serves.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[4].Text);
                }

                totrefs.Text = "" + dblSuma.ToString("n"+ deci);
                cantadd.Text = "1.0";
                //Limpar Arts de los textboxs
                //acant.Text = "1"; adesc.Text = ""; apart.Text = ""; aprecio.Text = ""; sistrefadd.Text = ""; sistmanadd.Text = ""; cvesatpa.Text = "25101500";
            }
            else
            {
                MessageBox.Show("No se puede agregar este concepto sin llenar todos los campos");
            }
        }

        private void delserv_Click(object sender, EventArgs e)
        {
            int sele = 0;
            double totalizar, subbo;
            
            Double dblSuma = 0;
            if (idser.Text != "")
            {
            sele = sele + int.Parse(idser.Text);
          
                if (DialogResult.Yes == MessageBox.Show("¿Desea eliminar?", "--Eliminar servicio--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                this.serves.Items[sele].Remove();
                idser.Text = "";
                //Sumar los elementos del listbox
                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[5].Text);
                }


                foreach (ListViewItem item in serves.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[4].Text);
                }

                totrefs.Text = "" + dblSuma.ToString("n" + deci);
                /////////////////Eliminarlo
                delserv.Enabled = false;
                }
            }
            else
            { MessageBox.Show("Debe seleccionar un servicio para eliminarlo"); }
                idser.Text = "";
        }

        private void serves_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (serves.SelectedIndices.Count <= 0)
            {
                return;
            }
            int intselectedindex = serves.SelectedIndices[0];
            if (intselectedindex >= 0)
            {
                //extrae el contenido del campo: MessageBox.Show(addreflist.Items[intselectedindex].Text);
                idser.Text = "" + intselectedindex;
                delserv.Enabled = true;
            }
        }

        private void claveserv_TextChanged(object sender, EventArgs e)
        {
            precserv.Text = "";
            double precios = 0;
            SqlConnection con = new SqlConnection(config.cade);
            SqlCommand cmd = new SqlCommand("SELECT precio FROM gastos WHERE clave= '" + claveserv.Text + "' ORDER BY clave ASC;", con);
            con.Open();
            SqlDataAdapter dapaas = new SqlDataAdapter(cmd);
            DataTable dtpaas = new DataTable();
            dapaas.Fill(dtpaas);
            int cuentapaas = dtpaas.Rows.Count;
            if (cuentapaas > 0)
            {
                foreach (DataRow rowp in dtpaas.Rows)
                {
                    precios = Math.Round(double.Parse("" + rowp["precio"].ToString()), deci, MidpointRounding.AwayFromZero);
                    precserv.Text = "" + precios.ToString("n" + deci);
                }
            }
            else { precserv.Text = ""; }
            con.Close();
        }

        private void updt_Click(object sender, EventArgs e)
        {
            int index = 0;
            if ((clavedit.Text != "") && (cantedit.Text != "") && (containeradd.Text != "") && (palled.Text != ""))
            {
                string adesc, clave, price, umes = "", contain;
                clave = "" + clavedit.Text;
                double acant1 = 0, aprecio1 = 0, totalon = 0, pallets = 0;
                acant1 = Math.Round(double.Parse(cantedit.Text), 2, MidpointRounding.AwayFromZero);
                pallets = Math.Round(double.Parse(palled.Text), 2, MidpointRounding.AwayFromZero);
                SqlConnection con = new SqlConnection(config.cade);
                SqlCommand cmd = new SqlCommand("SELECT clave, descr, precio, linea, ume FROM artsipl WHERE clave= '" + clave + "' ORDER BY id ASC;", con);
                con.Open();
                SqlDataAdapter dapaas = new SqlDataAdapter(cmd);
                DataTable dtpaas = new DataTable();
                dapaas.Fill(dtpaas);
                int cuentapaas = dtpaas.Rows.Count;
                if (cuentapaas > 0)
                {
                    foreach (DataRow rowp in dtpaas.Rows)
                    {
                        aprecio1 = Math.Round(double.Parse("" + rowp["precio"]), deci, MidpointRounding.AwayFromZero);
                        umes = "" + rowp["ume"];
                    }
                }



                if (listpr.Text != "")
                {
                    cmd = new SqlCommand("SELECT l.id,l.nom,p.precio FROM listasipl as l INNER JOIN pricesipl as p ON p.list = l.id WHERE p.clave='" + clave + "' AND l.nom='" + listpr.Text + "';", con);
                    dapaas = new SqlDataAdapter(cmd);
                    dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    cuentapaas = dtpaas.Rows.Count;
                    if (cuentapaas > 0)
                    {
                        foreach (DataRow rowp in dtpaas.Rows)
                        {
                            aprecio1 = Math.Round(double.Parse("" + rowp["precio"]),deci, MidpointRounding.AwayFromZero);
                        }
                    }
                    dapaas.Dispose(); dtpaas.Dispose(); cmd.Dispose();
                }
                con.Close();

                totalon = aprecio1 * acant1;

                //cant clave       ume         punit                         importe                           container
                index = int.Parse("" + idadded.Text);
                addrows.Items[index].Text = "" + acant1.ToString("n2");
                addrows.Items[index].SubItems[1].Text = "" + clave;
                addrows.Items[index].SubItems[2].Text = "" + umes;
                addrows.Items[index].SubItems[4].Text = "" + aprecio1.ToString("n" + deci);
                addrows.Items[index].SubItems[5].Text = "" + totalon.ToString("n" + deci);
                addrows.Items[index].SubItems[6].Text = "" + containeradd.Text;
                addrows.Items[index].SubItems[8].Text = "" + pallets.ToString("n2");

                addrows.Items[index].SubItems[3].Text = "" + kgboxx.Text;
                addrows.Items[index].SubItems[7].Text = "" + kilose.Text;
                //Sumar los elementos del listbox
                Double dblSuma = 0;

                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[5].Text);
                }

                foreach (ListViewItem item in serves.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[4].Text);
                }

                totrefs.Text = "" + dblSuma.ToString("n" + deci);


                //obtener peso neto
                dblSuma = 0;
                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[7].Text);
                }
                neto.Text = "" + dblSuma.ToString("n2");
                dblSuma = 0;

                //obtener peso bruto
                foreach (ListViewItem item in addrows.Items)
                {
                    dblSuma += Convert.ToDouble(item.SubItems[8].Text) * Convert.ToDouble(palletkgs.Text);
                }
                dblSuma = dblSuma + Convert.ToDouble(neto.Text);
                bruto.Text = "" + dblSuma.ToString("n2");
                dblSuma = 0;

                //calcular los totales en peso de cada cont
                logcontainer.Text = "";
                foreach (object item in contains.Items)
                {
                    dblSuma = 0;
                    foreach (ListViewItem items in addrows.Items)
                    {
                        if (item.ToString() + "" == "" + items.SubItems[6].Text)
                        {
                            dblSuma += Convert.ToDouble(items.SubItems[7].Text);
                        }
                    }
                    if (dblSuma > 0)
                    {
                        logcontainer.AppendText("-" + item.ToString() + ": " + dblSuma.ToString("n2") + " kgs.\r\n");
                    }
                    else
                    {
                        logcontainer.AppendText("-" + item.ToString() + ": 0 kgs.\r\n");
                    }
                }


                //Limpar Arts de los textboxs
                cantedit.Text = "";
                clavedit.Text = "";
                idadded.Text = "";
                palled.Text = "";
                kilose.Text = "";
                kgboxx.Text = "";
            }
            else
            {
                MessageBox.Show("No se puede actualizar este concepto sin llenar todos los campos");
            }
        }

        private void cantedit_KeyPress(object sender, KeyPressEventArgs e)
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

        private void palled_KeyPress(object sender, KeyPressEventArgs e)
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

        private void addrows_Click(object sender, EventArgs e)
        {
            cantedit.Text = "";
            clavedit.Text = "";
            idadded.Text = "";
            palled.Text = "";
            kgboxx.Text = "";
            kilose.Text = "";
            if (addrows.SelectedIndices.Count <= 0)
            {
                return;
            }
            int intselectedindex = addrows.SelectedIndices[0];
            if (intselectedindex >= 0)
            {
                idadded.Text = "" + intselectedindex;
                //extrae el contenido del campo: MessageBox.Show(addreflist.Items[intselectedindex].Text);
                cantedit.Text = "" + addrows.Items[intselectedindex].Text;
                clavedit.Text = "" + addrows.Items[intselectedindex].SubItems[1].Text;
                cue.Text = "" + addrows.Items[intselectedindex].SubItems[4].Text;
                kgboxx.Text = "" + addrows.Items[intselectedindex].SubItems[3].Text;
                kilose.Text = "" + addrows.Items[intselectedindex].SubItems[7].Text;
                containeradd.Text = "" + addrows.Items[intselectedindex].SubItems[6].Text;
                palled.Text = "" + addrows.Items[intselectedindex].SubItems[8].Text;
                delref.Enabled = true;
            }
        }

        private void kgboxx_KeyPress(object sender, KeyPressEventArgs e)
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

        private void kilose_KeyPress(object sender, KeyPressEventArgs e)
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

        private void serves_Click(object sender, EventArgs e)
        {
            if (serves.SelectedIndices.Count <= 0)
            {
                return;
            }
            int intselectedindex = serves.SelectedIndices[0];
            if (intselectedindex >= 0)
            {
                //extrae el contenido del campo: MessageBox.Show(addreflist.Items[intselectedindex].Text);
                idser.Text = "" + intselectedindex;
                delserv.Enabled = true;
            }
        }

        private void updt_Click_1(object sender, EventArgs e)
        {
            int index = 0;
            if (idadded.Text == "") { MessageBox.Show("Debe seleccionar una partida para editarla"); }
            else
            {
                if ((clavedit.Text != "") && (cantedit.Text != "") && (containeradd.Text != "") && (palled.Text != "") && (cue.Text != ""))
                {
                    if (DialogResult.Yes == MessageBox.Show("¿Desea actualizar?", "--Editar partida--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                    {
                        string adesc, clave, price, umes = "", contain;
                        clave = "" + clavedit.Text;
                        double acant1 = 0, aprecio1 = 0, totalon = 0, pallets = 0, cued = 0, kilosed = 0;
                        acant1 = Math.Round(double.Parse(cantedit.Text), 2, MidpointRounding.AwayFromZero);
                        pallets = Math.Round(double.Parse(palled.Text), 2, MidpointRounding.AwayFromZero);

                        cued = Math.Round(double.Parse(cue.Text), 2, MidpointRounding.AwayFromZero); ;
                        kilosed = Math.Round(double.Parse(kilose.Text), 2, MidpointRounding.AwayFromZero); ;


                        SqlConnection con = new SqlConnection(config.cade);
                        SqlCommand cmd = new SqlCommand("SELECT clave, descr, precio, linea, ume FROM artsipl WHERE clave= '" + clave + "' ORDER BY id ASC;", con);
                        con.Open();
                        SqlDataAdapter dapaas = new SqlDataAdapter(cmd);
                        DataTable dtpaas = new DataTable();
                        dapaas.Fill(dtpaas);
                        int cuentapaas = dtpaas.Rows.Count;
                        if (cuentapaas > 0)
                        {
                            foreach (DataRow rowp in dtpaas.Rows)
                            {
                                aprecio1 = Math.Round(double.Parse("" + rowp["precio"]), deci, MidpointRounding.AwayFromZero);
                                umes = "" + rowp["ume"];
                            }
                        }


                        /*
                        if (listpr.Text != "")
                        {
                            cmd = new SqlCommand("SELECT l.id,l.nom,p.precio FROM listasipl as l INNER JOIN pricesipl as p ON p.list = l.id WHERE p.clave='" + clave + "' AND l.nom='" + listpr.Text + "';", con);
                            dapaas = new SqlDataAdapter(cmd);
                            dtpaas = new DataTable();
                            dapaas.Fill(dtpaas);
                            cuentapaas = dtpaas.Rows.Count;
                            if (cuentapaas > 0)
                            {
                                foreach (DataRow rowp in dtpaas.Rows)
                                {
                                    aprecio1 = Math.Round(double.Parse("" + rowp["precio"]), 2, MidpointRounding.AwayFromZero);
                                }
                            }
                            dapaas.Dispose(); dtpaas.Dispose(); cmd.Dispose();
                        }
                        */
                        con.Close();
                        aprecio1 = Math.Round(double.Parse("" + cue.Text), deci, MidpointRounding.AwayFromZero);
                        totalon = aprecio1 * acant1;
                        //cant clave       ume         punit                         importe                           container
                        index = int.Parse("" + idadded.Text);
                        addrows.Items[index].Text = "" + acant1.ToString("n2");
                        addrows.Items[index].SubItems[1].Text = "" + clave;
                        addrows.Items[index].SubItems[2].Text = "" + umes;
                        addrows.Items[index].SubItems[4].Text = "" + aprecio1.ToString("n" + deci);
                        addrows.Items[index].SubItems[5].Text = "" + totalon.ToString("n" + deci);
                        addrows.Items[index].SubItems[6].Text = "" + containeradd.Text;
                        addrows.Items[index].SubItems[8].Text = "" + pallets.ToString("n2");

                        addrows.Items[index].SubItems[3].Text = "" + kgboxx.Text;
                        addrows.Items[index].SubItems[7].Text = "" + kilosed.ToString("n2");
                        //Sumar los elementos del listbox
                        Double dblSuma = 0;

                        foreach (ListViewItem item in addrows.Items)
                        {
                            dblSuma += Convert.ToDouble(item.SubItems[5].Text);
                        }

                        foreach (ListViewItem item in serves.Items)
                        {
                            dblSuma += Convert.ToDouble(item.SubItems[4].Text);
                        }

                        totrefs.Text = "" + dblSuma.ToString("n" + deci);


                        //obtener peso neto
                        dblSuma = 0;
                        foreach (ListViewItem item in addrows.Items)
                        {
                            dblSuma += Convert.ToDouble(item.SubItems[7].Text);
                        }
                        neto.Text = "" + dblSuma.ToString("n2");
                        dblSuma = 0;

                        //obtener peso bruto
                        foreach (ListViewItem item in addrows.Items)
                        {
                            dblSuma += Convert.ToDouble(item.SubItems[8].Text) * Convert.ToDouble(palletkgs.Text);
                        }
                        dblSuma = dblSuma + Convert.ToDouble(neto.Text);
                        bruto.Text = "" + dblSuma.ToString("n2");
                        dblSuma = 0;

                        //calcular los totales en peso de cada cont
                        logcontainer.Text = "";
                        foreach (object item in contains.Items)
                        {
                            dblSuma = 0;
                            foreach (ListViewItem items in addrows.Items)
                            {
                                if (item.ToString() + "" == "" + items.SubItems[6].Text)
                                {
                                    dblSuma += Convert.ToDouble(items.SubItems[7].Text);
                                }
                            }
                            if (dblSuma > 0)
                            {
                                logcontainer.AppendText("-" + item.ToString() + ": " + dblSuma.ToString("n2") + " kgs.\r\n");
                            }
                            else
                            {
                                logcontainer.AppendText("-" + item.ToString() + ": 0 kgs.\r\n");
                            }
                        }


                        //Limpar Arts de los textboxs
                        cantedit.Text = "";
                        clavedit.Text = "";
                        idadded.Text = "";
                        palled.Text = "";
                        kilose.Text = "";
                        kgboxx.Text = "";
                        cue.Text = "";
                    }
                }
                else
                {
                    MessageBox.Show("No se puede actualizar este concepto sin llenar todos los campos");
                }
            }
            idadded.Text = "";
        }

        private void cue_KeyPress(object sender, KeyPressEventArgs e)
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

        private void cantedit_KeyPress_1(object sender, KeyPressEventArgs e)
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

        private void kilose_KeyPress_1(object sender, KeyPressEventArgs e)
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

        private void palled_KeyPress_1(object sender, KeyPressEventArgs e)
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

        private void deces_ValueChanged(object sender, EventArgs e)
        {
            deci = Int32.Parse("" + deces.Value);
        }
    }
}
