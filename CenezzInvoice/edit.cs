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
    public partial class editor : Form
    {
        public int deci = 0;
        public editor()
        {
            InitializeComponent();
        }
        public string updated = "";

        private void editor_Load(object sender, EventArgs e)
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


            SqlDataReader reader = cmd.ExecuteReader();
            AutoCompleteStringCollection MyCollection = new AutoCompleteStringCollection();
            while (reader.Read())
            {
                MyCollection.Add(reader.GetString(0));
            }
            cveadd.AutoCompleteCustomSource = MyCollection;
            clavedit.AutoCompleteCustomSource = MyCollection;
            cmd.Dispose();reader.Dispose();

            
            cmd = new SqlCommand("SELECT clave FROM gastos ORDER BY clave ASC;", con);
            reader = cmd.ExecuteReader();
            AutoCompleteStringCollection MyCollectionS = new AutoCompleteStringCollection();
            while (reader.Read())
            {
                MyCollectionS.Add(reader.GetString(0));
            }
            claveserv.AutoCompleteCustomSource = MyCollectionS;
            cmd.Dispose(); reader.Dispose();



            query = "SELECT * FROM invoicespl WHERE id =" + config.idinvoice + ";";
            SqlCommand cm = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(cm);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    recnom.Text = "" + row["nomcli"];
                    dirrec.Text = "" + row["callecli"];
                    nume.Text = "" + row["numcli"];
                    numi.Text = "" + row["numclii"];
                    col.Text = "" + row["colcli"];
                    //cd.Text = "" + row["muncli"];
                    mun.Text = "" + row["muncli"];
                    edo.Text = "" + row["edocli"];
                    pais.Text = "" + row["paiscli"];
                   
                    emit.Text = "" + row["empresa"];
                    fecha.Text = "" + row["fecha"];
                    alba.Text = "" + row["albaran"];
                    folio.Text = "" + row["folio"];
                    oridest.Text = "" + row["origdest"];
                    totrefs.Text = "" + row["tot"];

                    obs1.Text = "" + row["obs1"];
                    obs2.Text = "" + row["obs2"];
                    obs3.Text = "" + row["obs3"];
                    obs4.Text = "" + row["obs4"];
                    obs5.Text = "" + row["obs5"];
                    foliaje.Text = "" + row["number"];
                    recep.Text = "" + row["idcli"];
                    neto.Text = "" + row["pesoneto"];
                    bruto.Text = "" + row["pesobruto"];
                    states.Text = "" + row["stats"];
                    if(row["stats"].ToString()=="CANCELADA")
                    {
                        states.ForeColor = Color.FromArgb(170, 0, 0);
                        button5.Enabled = false;
                        button9.Enabled = false;
                        button8.Enabled = false;
                        button2.Enabled = false;
                    }
                    else
                    {
                        states.ForeColor = Color.FromArgb(0, 170, 0);
                    }

                    currency.SelectedIndex = currency.FindStringExact("" + row["currency"]);
                    this.Text = "Editando Invoice # " + row["folio"];

                }
            }
            else
            {
                recep.Text = "";
                MessageBox.Show("Orden no existe");
                this.Close();
            }
            cm.Dispose();da.Dispose();dt.Dispose();

            listpr.Items.Add("");
            cmd = new SqlCommand("SELECT nom FROM listasipl ORDER BY nom ASC;", con);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                listpr.Items.Add(reader.GetString(0));
            }



            if (recep.Text != "")
            {
                niff.Text = "";
                query = "SELECT c.id, c.nom, c.nif,c.cd,c.precios,l.nom as nompre FROM clientesipl AS c INNER JOIN    listasipl as l ON l.id=c.precios WHERE c.id =" + recep.Text + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        niff.Text = "" + row["nif"];
                        cd.Text = "" + row["cd"];
                        listpr.Text = "" + row["nompre"];
                    }
                }
                else
                {
                    recep.Text = "";
                    MessageBox.Show("Cliente no existe");
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
            }

            if (emit.Text != "")
            {

                query = "SELECT nom FROM empresasipl WHERE id =" + emit.Text + ";";
                cm = new SqlCommand(query, con);
                da = new SqlDataAdapter(cm);
                dt = new DataTable();
                da.Fill(dt);
                cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        nomemp.Text = "" + row["nom"];
                    }
                }
                else
                {
                    emit.Text = "";
                    MessageBox.Show("Empresa no existe");
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();
            }



            query = "SELECT * FROM containersipl WHERE ord='" + config.idinvoice + "' ORDER BY id ASC;";
            
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    containeradd.Items.Add("" + row["container"]);
                    contains.Items.Add("" + row["container"]);
                    logcontainer.AppendText("-" + row["container"] + ": " + row["pesoneto"] + " kgs.\r\n");
                }
            }
            cm.Dispose(); da.Dispose(); dt.Dispose();




            query = "SELECT r.cant, r.clave, r.ume, r.pu, r.importe, r.container, r.pallets,r.pesoneto, r.pesobruto,ISNULL((SELECT kgscaja FROM artsipl WHERE clave=r.clave),0) AS kcaja,r.cajas  FROM rowsipl AS r WHERE r.ord='" + config.idinvoice + "' ORDER BY r.id ASC;";
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    //cant, clave, ume, kilos caja, p unitario, importe, container, kgs tot, pallets
                    string[] row1 = { "" + row["clave"], "" + row["ume"], "" + row["kcaja"], "" + row["pu"], "" + row["importe"], "" + row["container"], "" + row["pesoneto"], "" + row["pallets"], "" + row["cajas"] };
                    addrows.Items.Add("" + row["cant"]).SubItems.AddRange(row1);
                }
            }
            cm.Dispose(); da.Dispose(); dt.Dispose();




            query = "SELECT cant, clave, descrip, cu, total FROM  rowsservpl WHERE ord ='" + config.idinvoice + "' ORDER BY id ASC;";
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    string[] row1 = { "" + row["clave"], "" + row["descrip"], "" + row["cu"], "" + row["total"] };
                    serves.Items.Add("" + row["cant"]).SubItems.AddRange(row1);
                }
            }
            cm.Dispose(); da.Dispose(); dt.Dispose();


            deci = Int32.Parse("" + deces.Value);
            con.Close();
        }

        private void editor_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if ((cveadd.Text != "") && (cantadd.Text != "") && (containeradd.Text != ""))
            {
                string adesc, clave, price, umes = "", contain, pallet = "";
                clave = "" + cveadd.Text;
                double acant1 = 0, aprecio1 = 0, kgspiece = 0, totalon = 0, pallets = 0, kgscaja = 0, kilos = 0, metros = 0, cajasp = 0, cajasm2 = 0;


                acant1 = Math.Round(double.Parse(cantadd.Text), 2, MidpointRounding.AwayFromZero);

                SqlConnection con = new SqlConnection(config.cade);
                SqlCommand cmd = new SqlCommand("SELECT clave,descr,precio,linea,ume,pallet,kgspiece,kgscaja,mtscaja,caja FROM artsipl WHERE clave= '" + clave + "' ORDER BY clave ASC;", con);
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
                        cajasp = Math.Round(double.Parse("" + rowp["caja"]), 2, MidpointRounding.AwayFromZero);
                        cajasm2 = Math.Round(double.Parse("" + rowp["mtscaja"]), 2, MidpointRounding.AwayFromZero);
                    }
                }

                dapaas.Dispose(); dtpaas.Dispose(); cmd.Dispose();

                if (listpr.Text != "")
                {
                    cmd = new SqlCommand("SELECT l.id,l.nom,p.precio FROM listasipl as l INNER JOIN pricesipl as p ON p.list = l.id WHERE p.clave='" + clave + "' AND l.nom='" + listpr.Text + "';", con);
                    //MessageBox.Show("" + "SELECT l.id,l.nom,p.precio FROM listasipl as l INNER JOIN pricesipl as p ON p.list = l.id WHERE p.clave='" + clave + "' AND l.nom='" + listpr.Text + "';");
                    dapaas = new SqlDataAdapter(cmd);
                    dtpaas = new DataTable();
                    dapaas.Fill(dtpaas);
                    cuentapaas = dtpaas.Rows.Count;
                    if (cuentapaas > 0)
                    {
                        foreach (DataRow rowp in dtpaas.Rows)
                        {
                            aprecio1 = Math.Round(double.Parse("" + rowp["precio"]), deci, MidpointRounding.AwayFromZero);
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

                if (umes == "m²")
                {
                    try
                    {
                        cajasm2 = double.Parse("" + acant1) / metros;
                    }
                    catch { cajasm2 = 1; }
                    //if (pallets < 1) { pallets = 1; }
                }
                else
                {
                    try
                    {
                        cajasm2 = double.Parse("" + acant1) / cajasp;
                    }
                    catch { cajasm2 = 1; }

                }

                if (cajasm2.ToString() == "∞") { cajasm2 = 0; }

                try
                {
                    kilos = double.Parse("" + cajasm2) * kgscaja;
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
                string[] row1 = { "" + clave, "" + umes, kgscaja.ToString("n2"), "" + aprecio1.ToString("n" + deci), "" + totalon.ToString("n" + +deci), "" + containeradd.Text, "" + kilos.ToString("n2"), "" + pallets.ToString("n2"), "" + cajasm2.ToString("n2") };
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
                neto.Text = "" + dblSuma.ToString("n" + deci);
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
                        logcontainer.AppendText("-" + item.ToString() + ": " + dblSuma.ToString("n2") + " kgs.\r\n");
                    }
                }


                cantadd.Text = "1.0";
                //Limpar Arts de los textboxs
                //acant.Text = "1"; adesc.Text = ""; apart.Text = ""; aprecio.Text = ""; sistrefadd.Text = ""; sistmanadd.Text = ""; cvesatpa.Text = "25101500";
            }
            else
            {
                MessageBox.Show("No se puede agregar este concepto sin llenar todos los campos");
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
                foreach (object item in contains.Items)
                {
                    itac = "" + item.ToString();
                    if (itac == "" + toadd)
                    {
                        addded = "no";
                    }
                }

                if (addded == "yes")
                {
                    containsadd.Text = "";
                    contains.Items.Add("" + toadd);
                    containeradd.Items.Add("" + toadd);

                    string qui = "INSERT INTO containersipl (ord,container) VALUES ('" + config.idinvoice + "','" + toadd + "');";
                    SqlCommand myCo = new SqlCommand(qui, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    containeradd.Items.Clear();
                    foreach (object item in contains.Items)
                    {
                            itac = "" + item.ToString();
                            containeradd.Items.Add("" + itac);
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

                string qu = "DELETE FROM containersipl WHERE ord = '"+config.idinvoice + "' AND container = '"+ selectedItem  + "';";
                MessageBox.Show(" dele: " + qu );
                SqlCommand myCommandref = new SqlCommand(qu, config.conn);
                myCommandref.ExecuteNonQuery();
                myCommandref.Dispose();

                containeradd.Items.Clear();
                contains.Items.Clear();
                qu = "SELECT container,precinto FROM containersipl WHERE ord='" + config.idinvoice + "' ORDER BY id ASC;";

                SqlCommand cm = new SqlCommand(qu, config.conn);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        containeradd.Items.Add("" + row["container"]);
                        contains.Items.Add("" + row["container"]);
                    }
                }
                cm.Dispose(); da.Dispose(); dt.Dispose();



                foreach (ListViewItem item in addrows.Items)
                {
                    if (selectedItem == "" + item.SubItems[5].Text)
                    {
                        addrows.Items[countitems].SubItems[5].Text = "" ;
                    }
                    countitems = countitems + 1;
                }
            }
            else
            {
                MessageBox.Show("Seleccione un contenedor de la lista para quitarlo.");
            }
        }

        private void delref_Click(object sender, EventArgs e)
        {
 
            int sele = 0;
            double totalizar, subbo;
            sele = sele + int.Parse(idadded.Text);
            Double dblSuma = 0;
            if (idadded.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea eliminar?", "--Eliminar partida--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    addrows.Items[sele].Remove();
                    idadded.Text = "";
                    //totalizar monto final en moneda
                    foreach (ListViewItem item in addrows.Items)
                    {
                        dblSuma += Convert.ToDouble(item.SubItems[5].Text);
                    }

                    foreach (ListViewItem item in serves.Items)
                    {
                        dblSuma += Convert.ToDouble(item.SubItems[4].Text);
                    }
                    totrefs.Text = "" + dblSuma.ToString("n2");


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


                    /////////////////Eliminarlo
                    cveadd.Focus();
                    delref.Enabled = false;
                }
            }
            else { MessageBox.Show("Debe seleccionar una partida para eliminar"); }
            idadded.Text = "";
            }

        private void addrows_SelectedIndexChanged(object sender, EventArgs e)
        {
            cantedit.Text = "";
            clavedit.Text = "";
            idadded.Text = "";
            palled.Text = "";
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
                //listpr.Text = "" + addrows.Items[intselectedindex].SubItems[1].Text;
                containeradd.Text = "" + addrows.Items[intselectedindex].SubItems[5].Text;
                palled.Text = "" + addrows.Items[intselectedindex].SubItems[6].Text;
                delref.Enabled = true;
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
                string query = "SELECT id, nom, nif, calle, num, numi, col, mun, cd, edo, pais, cp FROM clientesipl WHERE id =" + recep.Text + ";";
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
                    }
                }
                else
                {
                    recep.Text = "";
                    MessageBox.Show("Empresa no existe");
                }
                da.Dispose(); cm.Dispose(); dt.Dispose();
                con.Close();

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
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
                        try
                        {
                            string qu = "UPDATE  invoicespl SET folio='" + folio.Text + "', empresa='" + emit.Text + "', nomcli='" + recnom.Text + "', callecli='" + dirrec.Text + "', numcli='" + nume.Text + "', numclii='" + numi.Text + "', colcli='" + col.Text + "', muncli='" + mun.Text + "', edocli='" + edo.Text + "', paiscli='" + pais.Text + "', fecha='" + fecha.Text + "', albaran='" + alba.Text + "', origdest='" + oridest.Text + "',tot='" + totrefs.Text + "',idcli='" + recep.Text + "',currency='" + currency.Text + "',obs1='" + obs1.Text + "',obs2='" + obs2.Text + "',obs3='" + obs3.Text + "',obs4='" + obs4.Text + "',obs5='" + obs5.Text + "', pesoneto='" + neto.Text + "',pesobruto='" + bruto.Text + "' WHERE id=" + config.idinvoice;
                            SqlCommand myCo = new SqlCommand(qu, config.conn);
                            myCo.ExecuteNonQuery();
                            myCo.Dispose();

                            qu = "DELETE FROM rowsipl WHERE ord='" + config.idinvoice + "';";
                            myCo = new SqlCommand(qu, config.conn);
                            myCo.ExecuteNonQuery();
                            myCo.Dispose();

                            qu = "DELETE FROM rowsservpl WHERE ord='" + config.idinvoice + "';";
                            myCo = new SqlCommand(qu, config.conn);
                            myCo.ExecuteNonQuery();
                            myCo.Dispose();

                            /*
                            qu = "DELETE FROM containersipl WHERE ord='" + config.idinvoice + "';";
                            myCo = new SqlCommand(qu, config.conn);
                            myCo.ExecuteNonQuery();
                            myCo.Dispose();
                            */

                            string itac = ""; double dblSuma = 0, brutocont = 0;
                            foreach (object item in contains.Items)
                            {
                                itac = "" + item.ToString();
                                if (itac != "")
                                {
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
                                        qu = "UPDATE containersipl SET pesoneto='" + dblSuma.ToString("n2") + "', pesobruto='" + brutocont.ToString("n2") + "' WHERE container= '" + itac + "' AND  ord='" + config.idinvoice + "';";
                                        myCo = new SqlCommand(qu, config.conn);
                                        myCo.ExecuteNonQuery();
                                        myCo.Dispose();
                                    }
                                    //fin de kilos en container
                                }
                            }

                            //insertar partidas de la factura
                            string queryinsref = ""; brutocont = 0;
                            foreach (ListViewItem item in addrows.Items)
                            {                                    // 0    1    2      4      5      6         7     8
                                                                 //cant,clave,ume, punit,importe,container,kgstot,pallets
                                brutocont = (Convert.ToDouble("" + palletkgs.Text) * Convert.ToDouble("" + item.SubItems[8].Text)) + Convert.ToDouble(item.SubItems[7].Text);
                                queryinsref = "INSERT INTO rowsipl ( ord, cant, clave, ume, pu, importe, container,pallets,pesoneto, pesobruto,cajas)" +
                                " VALUES ('" + config.idinvoice + "', '" + item.SubItems[0].Text + "', '" + item.SubItems[1].Text + "', '" + item.SubItems[2].Text + "', '" + item.SubItems[4].Text + "', '" + item.SubItems[5].Text + "', '" + item.SubItems[6].Text + "','" + item.SubItems[8].Text + "','" + item.SubItems[7].Text + "','" + brutocont.ToString("n2") + "','" + item.SubItems[9].Text + "');";
                                // 0                                      1                                      2                              4                            5                                      6                        7                                  8
                                //cant,                            clave,                                 ume,                         punit,                        importe,                              container,                  kgstot,                          pallets
                                SqlCommand myCommandref = new SqlCommand(queryinsref, config.conn);
                                myCommandref.ExecuteNonQuery();
                                myCommandref.Dispose();
                            }
                            //insertar servicios
                            queryinsref = "";
                            foreach (ListViewItem item in serves.Items)
                            {
                                queryinsref = "INSERT INTO rowsservpl (  ord, cant, clave, descrip, cu,  total) VALUES ('" + config.idinvoice + "', '" + item.SubItems[0].Text + "', '" + item.SubItems[1].Text + "', '" + item.SubItems[2].Text + "', '" + item.SubItems[3].Text + "', '" + item.SubItems[4].Text + "');";
                                SqlCommand myCommandref = new SqlCommand(queryinsref, config.conn);
                                myCommandref.ExecuteNonQuery();
                                myCommandref.Dispose();
                            }

                            //list.FillDataGridView();
                            updated = "YES";
                            this.Close();

                        }
                        catch (Exception ex) { MessageBox.Show("Error al actualizar invoice.\r\n" + ex.ToString()); }
                    }
            }
        }
        }

        private void contains_SelectedIndexChanged(object sender, EventArgs e)
        {
            preci.Text = "";
            updtc.Text = "";nucon.Text = "";
            if (contains.SelectedIndices.Count <= 0)
            {
                return;
            }
            int intselectedindex = contains.SelectedIndices[0];
            if (intselectedindex >= 0)
            {
                //extrae el contenido del campo: MessageBox.Show(addreflist.Items[intselectedindex].Text);
                //updtc.Text = "" + intselectedindex;
                nucon.Text = "" + intselectedindex;
                updtc.Text = "" + contains.Items[intselectedindex].ToString();

                string query = "SELECT precinto FROM containersipl WHERE ord='" + config.idinvoice + "' AND container = '"+ contains.Items[intselectedindex].ToString() + "';";
                SqlCommand cm = new SqlCommand(query, config.conn);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        preci.Text = "" + row["precinto"];
                    }
                }
                else { preci.Text = "SIN PRECINTO"; }
                cm.Dispose(); da.Dispose(); dt.Dispose();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (nucon.Text == "")
            {
                MessageBox.Show("No se puede actualizar sin un numero de contenedor seleccionado.");
            }
            else
            {
                string anterior = "",anteriorrecinto="";
                if (DialogResult.Yes == MessageBox.Show("¿Desea actualizar?\r\nNuevo numero de contenedor: " + updtc.Text + ".", "--Cambio de número de contenedor--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    if (contains.SelectedIndices.Count <= 0)
                    {
                        MessageBox.Show("No se puede actualizar sin un numero de contenedor seleccionado.");
                    }
                    else
                    {
                        int intselectedindex = contains.SelectedIndices[0];
                        if (intselectedindex >= 0)
                        {
                            anterior = "" + contains.Items[intselectedindex].ToString();
                            anteriorrecinto = "" + preci.Text;
                            MessageBox.Show("Actualizando el contenedor:\r\n" + anterior + " por: " + updtc.Text + "", "Cambio de numero de contenedor", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            contains.Items[intselectedindex] = "" + updtc.Text;

                            int iC = containeradd.FindStringExact("" + anterior);
                            if (iC >= 0)
                            {
                                containeradd.Items[iC] = "" + updtc.Text;
                            }
                            int countitems = 0; string queryinsref = "";
                            SqlCommand myCommandref = new SqlCommand("", config.conn);
                            myCommandref.Dispose();

                            foreach (ListViewItem item in addrows.Items)
                            {
                                if (anterior == "" + item.SubItems[5].Text)
                                {
                                    addrows.Items[countitems].SubItems[5].Text = "" + updtc.Text;
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


                            queryinsref = "UPDATE rowsipl SET container ='" + updtc.Text + "' WHERE ord ='" + config.idinvoice + "' AND container ='" + anterior + "';";
                            myCommandref = new SqlCommand(queryinsref, config.conn);
                            myCommandref.ExecuteNonQuery();
                            myCommandref.Dispose();

                            queryinsref = "UPDATE containersipl SET container ='" + updtc.Text + "',precinto='" + anteriorrecinto + "' WHERE ord ='" + config.idinvoice + "' AND container ='" + anterior + "';";
                            myCommandref = new SqlCommand(queryinsref, config.conn);
                            myCommandref.ExecuteNonQuery();
                            myCommandref.Dispose();

                            containeradd.Items.Clear();
                            foreach (object item in contains.Items)
                            {
                                containeradd.Items.Add("" + item.ToString());
                            }
                        }
                        updtc.Text = ""; preci.Text = "";
                    }

                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string sclave;
            string descripte = "";
            if ((claveserv.Text != "") && (cantserv.Text != "") && (precserv.Text != ""))
            {
                sclave = "" + claveserv.Text;
                double cants = 0, precios = 0, totalon = 0;
                cants = Math.Round(double.Parse(cantserv.Text), 2, MidpointRounding.AwayFromZero);
                precios = Math.Round(double.Parse(precserv.Text), 2, MidpointRounding.AwayFromZero);

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
                string[] row1 = { "" + sclave, "" + descripte, "" + precios.ToString("n2"), "" + totalon.ToString("n2") };
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

                totrefs.Text = "" + dblSuma.ToString("n2");
                cantadd.Text = "1.0";
                //Limpar Arts de los textboxs
                //acant.Text = "1"; adesc.Text = ""; apart.Text = ""; aprecio.Text = ""; sistrefadd.Text = ""; sistmanadd.Text = ""; cvesatpa.Text = "25101500";
            }
            else
            {
                MessageBox.Show("No se puede agregar este concepto sin llenar todos los campos");
            }

        }

        private void cantadd_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
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

        private void button3_Click(object sender, EventArgs e)
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

                    totrefs.Text = "" + dblSuma.ToString("n2");

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
                    precios = Math.Round(double.Parse("" + rowp["precio"].ToString()), 2, MidpointRounding.AwayFromZero);
                    precserv.Text = "" + precios.ToString("n2");
                }
            }
            else { precserv.Text = ""; }
            con.Close();
        }

        private void updt_Click(object sender, EventArgs e)
        {

            int index = 0;
            if (idadded.Text == "") { MessageBox.Show("Debe seleccionar una partida para editarla");  }
            else
            {
                if ((clavedit.Text != "") && (cantedit.Text != "") && (containeradd.Text != "") && (palled.Text != "") && (cue.Text != ""))
                {
                    if (DialogResult.Yes == MessageBox.Show("¿Desea actualizar?", "--Editar partida--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                    {
                        string adesc, clave, price, umes = "", contain;
                        clave = "" + clavedit.Text;
                        double acant1 = 0, aprecio1 = 0, totalon = 0, pallets = 0, cued = 0, kilosed = 0, cajaded=0;
                        acant1 = Math.Round(double.Parse(cantedit.Text), 2, MidpointRounding.AwayFromZero);
                        pallets = Math.Round(double.Parse(palled.Text), 2, MidpointRounding.AwayFromZero);

                        cued = Math.Round(double.Parse(cue.Text), 2, MidpointRounding.AwayFromZero); ;
                        kilosed = Math.Round(double.Parse(kilose.Text), 2, MidpointRounding.AwayFromZero); ;
                        cajaded = Math.Round(double.Parse(cajass.Text), 2, MidpointRounding.AwayFromZero); ;

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
                                aprecio1 = Math.Round(double.Parse("" + rowp["precio"]), 2, MidpointRounding.AwayFromZero);
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
                        aprecio1 = Math.Round(double.Parse("" + cue.Text), 2, MidpointRounding.AwayFromZero);
                        totalon = aprecio1 * acant1;
                        //cant clave       ume         punit                         importe                           container
                        index = int.Parse("" + idadded.Text);
                        addrows.Items[index].Text = "" + acant1.ToString("n2");
                        addrows.Items[index].SubItems[1].Text = "" + clave;
                        addrows.Items[index].SubItems[2].Text = "" + umes;
                        addrows.Items[index].SubItems[4].Text = "" + aprecio1.ToString("n2");
                        addrows.Items[index].SubItems[5].Text = "" + totalon.ToString("n2");
                        addrows.Items[index].SubItems[6].Text = "" + containeradd.Text;
                        addrows.Items[index].SubItems[8].Text = "" + pallets.ToString("n2");
                        addrows.Items[index].SubItems[9].Text = "" + cajaded.ToString("n2");

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

                        totrefs.Text = "" + dblSuma.ToString("n2");


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
                        cajass.Text = "";
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

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void clavedit_TextChanged(object sender, EventArgs e)
        {

        }

        private void cantedit_TextChanged(object sender, EventArgs e)
        {

        }

        private void palled_TextChanged(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void addrows_Click(object sender, EventArgs e)
        {
            cantedit.Text = "";
            clavedit.Text = "";
            idadded.Text = "";
            palled.Text = "";
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
                //listpr.Text = "" + addrows.Items[intselectedindex].SubItems[1].Text;
                containeradd.Text = "" + addrows.Items[intselectedindex].SubItems[5].Text;
                palled.Text = "" + addrows.Items[intselectedindex].SubItems[6].Text;
                delref.Enabled = true;
            }
        }

        private void editor_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(updated == "YES") {

            }
    }

        private void addrows_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            cantedit.Text = "";
            clavedit.Text = "";
            idadded.Text = "";
            palled.Text = "";
            kgboxx.Text = "";
            kilose.Text = "";
            cajass.Text = "";
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
                cajass.Text = "" + addrows.Items[intselectedindex].SubItems[9].Text;
                delref.Enabled = true;
            }
        }

        private void addrows_Click_1(object sender, EventArgs e)
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

        private void serves_Click(object sender, EventArgs e)
        {

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
    }
}
