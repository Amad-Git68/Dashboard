using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using iTextSharp;
using iTextSharp.text.pdf;
using Ionic.Zip;
using System.Text; //@tri2

public partial class FSC_MRILetter : System.Web.UI.Page
{
    private string ConString = Main.ConnString.ConnStringSISD();
    private SqlDatabase con;

    private bool isView;
    private bool isAdd;
    private bool isDelete;
    private bool isEdit;

    protected void Page_Load(object sender, EventArgs e)
    {
        var strUserLogin = Session["userName"] == null ? "" : Session["userName"].ToString();

        var strPriv = Main.FunctionCollection.GetPrivilege("MRI_LETTER", strUserLogin);
        isView = strPriv.Substring(0, 1) == "1";

        if (!isView)
        {
            //Log
            var strAct = "Unauthorized user tried to access this menu";
            AddLog(strUserLogin, strAct);

            Response.Redirect("~/Default.aspx");
        }

        if (!Page.IsPostBack)
        {
            BindLetter();
            Clear();

            lblDate.Text = DateTime.Now.ToString("dd-MM-yyyy");

        }
    }

    private void BindLetter()
    {
        try
        {
            // ConString = "Data Source=SQLSUBSYSTEM-DB\\SQLSUBSYSTEMDB;Initial Catalog=SequisIS;User ID=developer;Password=P@ssw0rd2016;MultipleActiveResultSets=true;Connection Timeout=600; ";
            var strUser = Session["userName"].ToString();

            con = new SqlDatabase(ConString);

            var strSql = "select distinct letter_type," +
                         "  case letter_type" +
                         "      when 'SA' then 'Surat Akseptasi'" +
                         "      when 'SE' then 'Surat Ekstra Premi'" +
                         "      when 'SR' then 'Surat Reject'" +
                         "      when 'SM' then 'Surat Medical'" +
                         "      when 'SP' then 'Surat Postpone'" +
                         "      when 'SS' then 'Surat Surender'" +
                         "      when 'SC' then 'Surat Claim'" +
                         "      when 'SD' then 'Surat Decline'" +
                         "      when 'WS' then 'Worksheet'" +
                         "      when 'PL' then 'Invoice PL Niaga'" +
                         "      when 'NS' then 'Invoice KTA OCBC NISP'" +           //***amad***
                         "      when 'LP' then 'List Peserta'" +                    //List peserta 
                         "      when 'SB' then 'Surat BPP'" +
                         "      when 'AJ' then 'Surat AJK Restruktur'"+
                         "      else 'Billing/Debit Note'" +
                         "  end letter_desc " +
                         "from tbl_fsc_mri_letter where login_name='" + strUser + "' " +
                         "order by letter_type";

            var dt = con.ExecuteQuery(strSql);
            if (dt != null)
            {
                ddlLetter.DataSource = dt;
                ddlLetter.DataValueField = "letter_type";
                ddlLetter.DataTextField = "letter_desc";
                ddlLetter.DataBind();
            }

            ddlLetter.Items.Insert(0, new ListItem("-- Select Letter --", "0"));
        }
        catch (Exception ex)
        {
            ShowAlert(ex.Message + "Please contact Administrator.", "warn");
        }
    }

    protected void btnProses_Click(object sender, EventArgs e)
    {
        try
        {
            var strLetter = ddlLetter.SelectedValue;
            hdnLetter.Value = strLetter;
            var strUser = Session["userName"].ToString();
      //    var strLet  = Session["strLetter"].ToString();

            var arrConfig = Main.Configuration.GetConfigArnet();
            var arrServer = arrConfig[0].Split('/');
            var strUname = arrConfig[1];
            var strPass = arrConfig[2];
            var strServer = arrServer[2];

            con = new SqlDatabase(ConString);
            var strSql = "select file_name,path from tbl_fsc_mri_letter where letter_type='" + strLetter + "' and login_name='" + strUser + "'";
            var dr = con.GetOneRow(strSql);

            var strFname = dr["file_name"].ToString(); //"dnletand.dbf";
            var strPath = dr["path"].ToString(); //"//u/temp/";
            var strFileLoc = "ftp://" + strServer + "/" + strPath + strFname;
            var strLocalFile = Main.ServerPath() + "\\TempFile\\Arnet\\" + strFname;
            Session["nmfile"] = strFname;

            var strResult = DownloadFile(strFileLoc, strLocalFile, strUname, strPass);
            if (strResult != "Download Success")
                ShowAlert("Failed when downloading data from FTP. " + strResult, "warn");
            else
                ShowDataGrid(strFname);

            if (!chkZIP.Checked)
            {
                btnDownload_Click(new object(), new EventArgs()); //@tri
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex);
            ShowAlert("Failed. " + ex.Message, "warn");
        }
    }

    private void ShowDataGrid(string fileName)
    {
        var strPathDB = Main.ServerPath() + "\\TempFile\\Arnet\\";
        var strFileLoc = strPathDB + fileName;

        gvData.Columns.Clear();
        gvData.DataBind();

        if (File.Exists(strFileLoc))
        {
            var strConnstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strPathDB + ";Extended Properties=dBASE IV;";
            try
            {
                OleDbConnection dbCon = new OleDbConnection(strConnstring);
                dbCon.Open();

                var strSql = "SELECT * FROM " + fileName;
                OleDbDataAdapter dbDA = new OleDbDataAdapter(strSql, dbCon);
                DataTable dt = new DataTable();
                dbDA.Fill(dt);

                DataTable dtCust = new DataTable();

                string[] arrColumn;
                string[] arrHeader;

                if (fileName.IndexOf("dnlet") > -1)
                {
                    arrColumn = new string[] { "no", "npol", "nam", "prm", "up", "ctr", "desc", "alamat", "kota" };
                    arrHeader = new string[] { "No", "Policy No.", "Tertanggung", "Premi", "UP", "Masa Asuransi", "Desc", "Alamat", "Kota" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = dt.Rows[i]["nam"].ToString();
                        //arrData[3] = Convert.ToDateTime(dt.Rows[i]["dob"].ToString()).ToString("dd-MM-yyyy");
                        arrData[3] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[4] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[5] = dt.Rows[i]["ctr"].ToString();
                        arrData[6] = dt.Rows[i]["alamat1"].ToString() == "" ? dt.Rows[i]["desc"].ToString() : dt.Rows[i]["alamat1"].ToString();
                        arrData[7] = dt.Rows[i]["alamat"].ToString() == "" ? dt.Rows[i]["alamat1"].ToString() : dt.Rows[i]["alamat"].ToString();
                        arrData[8] = dt.Rows[i]["kota"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = true;
                }
                else if (fileName.IndexOf("listcr") > -1)   //List peserta 
                {
                    arrColumn = new string[] { "no", "npol", "bcode", "npst", "nam", "prm", "desc" };
                    arrHeader = new string[] { "No", "npol", "kCabang", "No_peserta", "Nama", "premi", "Cabang" };
                    dtCust = PopulateColumn(arrColumn, arrHeader);
                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = dt.Rows[i]["bcode"].ToString();
                        arrData[3] = Convert.ToInt64(dt.Rows[i]["npst"]).ToString();
                        arrData[4] = dt.Rows[i]["nam"].ToString();
                        arrData[5] = Convert.ToInt64(dt.Rows[i]["prm"]).ToString();
                        arrData[6] = dt.Rows[i]["desc"].ToString();
                        dtCust.Rows.Add(arrData);
                        x++;
                    }
                    chkPDF.Visible = false;
                }
                else if (fileName.IndexOf("notaocbc") > -1)   //nota OCBC NISP     ***amad***
                {
                    arrColumn = new string[] { "no", "npol", "th", "bl", "prmgrs", "pdisc", "tprm" };
                    arrHeader = new string[] { "No", "npol", "thn", "bln", "Premi Gross", "Komisi Disc", "Netto Premi" };
                    dtCust = PopulateColumn(arrColumn, arrHeader);
                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = Convert.ToInt64(dt.Rows[i]["th"]).ToString();
                        arrData[3] = Convert.ToInt64(dt.Rows[i]["bl"]).ToString();
                        arrData[4] = Convert.ToInt64(dt.Rows[i]["prmgrs"]).ToString();
                        arrData[5] = Convert.ToInt64(dt.Rows[i]["pdisc"]).ToString();
                        arrData[6] = Convert.ToInt64(dt.Rows[i]["tprm"]).ToString();
                        dtCust.Rows.Add(arrData);
                        x++;
                    }
                    chkPDF.Visible = false;
                }
                else if (fileName.IndexOf("dnlt1") > -1)  //Decline
                {
                    arrColumn = new string[] { "no", "npol", "nam", "dob", "up", "ctr", "desc" };
                    arrHeader = new string[] { "No", "Policy No.", "Tertanggung", "Tanggal Lahir", "UP", "Masa Asuransi", "Desc" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = dt.Rows[i]["nam"].ToString();
                        arrData[3] = Convert.ToDateTime(dt.Rows[i]["dob"].ToString()).ToString("dd-MM-yyyy");
                        arrData[4] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[5] = dt.Rows[i]["ctr"].ToString();
                        arrData[6] = dt.Rows[i]["desc"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = false;
                }
                else if (fileName.IndexOf("wrksh") > -1)
                {
                    arrColumn = new string[] { "no", "npol", "nam", "dob", "prm", "up", "ctr", "desc" };
                    arrHeader = new string[] { "No", "Policy No.", "Tertanggung", "Tanggal Lahir", "Premi", "UP", "Masa Asuransi", "Desc" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = dt.Rows[i]["nam"].ToString();
                        arrData[3] = Convert.ToDateTime(dt.Rows[i]["dob"].ToString()).ToString("dd-MM-yyyy");
                        //arrData[4] = dt.Rows[i]["sex"].ToString() == "True"? "Male" : "Female";
                        arrData[4] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[5] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[6] = dt.Rows[i]["ctr"].ToString();
                        arrData[7] = dt.Rows[i]["desc"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = false;
                }
                else if (fileName.IndexOf("dnlt3") > -1)  //Extra
                {
                    arrColumn = new string[] { "no", "nam", "dob", "up", "prm", "eprm", "ctr", "desc" };
                    arrHeader = new string[] { "No", "Tertanggung", "Tanggal Lahir", "UP", "Premi Dasar", "Tambahan", "Masa Asuransi", "Desc" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["nam"].ToString();
                        arrData[2] = Convert.ToDateTime(dt.Rows[i]["dob"].ToString()).ToString("dd-MM-yyyy");
                        arrData[3] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[4] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[5] = "Rp. " + Convert.ToInt64(dt.Rows[i]["eprm"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[6] = dt.Rows[i]["ctr"].ToString();
                        arrData[7] = dt.Rows[i]["desc"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = true;
                }
                else if (fileName.IndexOf("dnlt4") > -1)   //medical
                {
                    arrColumn = new string[] { "no", "npol", "nam", "dob", "up", "ctr", "desc" };
                    arrHeader = new string[] { "No", "Policy No.", "Tertanggung", "Tanggal Lahir", "UP", "Masa Asuransi", "Desc" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = dt.Rows[i]["nam"].ToString();
                        arrData[3] = Convert.ToDateTime(dt.Rows[i]["dob"].ToString()).ToString("dd-MM-yyyy");
                        arrData[4] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[5] = dt.Rows[i]["ctr"].ToString();
                        arrData[6] = dt.Rows[i]["desc"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = false;
                }
                else if (fileName.IndexOf("dnlt5") > -1)  //Postpone
                {
                    arrColumn = new string[] { "no", "npol", "nam", "dob", "up", "ctr", "desc" };
                    arrHeader = new string[] { "No", "Policy No.", "Tertanggung", "Tanggal Lahir", "UP", "Masa Asuransi", "Desc" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = dt.Rows[i]["nam"].ToString();
                        arrData[3] = Convert.ToDateTime(dt.Rows[i]["dob"].ToString()).ToString("dd-MM-yyyy");
                        arrData[4] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[5] = dt.Rows[i]["ctr"].ToString();
                        arrData[6] = dt.Rows[i]["desc"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = false;
                }
                else if (fileName.IndexOf("dnlt6") > -1)  //surender
                {
                    arrColumn = new string[] { "no", "nam", "up", "prm", "Tgl_bayar", "prm_pay", "desc" };
                    arrHeader = new string[] { "No", "Tertanggung", "UP", "Premi Dasar", "Tgl.Surender", "Nilai Surender", "Bank" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["nam"].ToString();
                        arrData[2] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[3] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[4] = dt.Rows[i]["Tgl_bayar"].ToString();
                        arrData[5] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm_pay"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[6] = dt.Rows[i]["desc"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = true;
                }
                else if (fileName.IndexOf("dnlt7") > -1)   //claim
                {
                    arrColumn = new string[] { "no", "nam", "up", "prm", "Tgl_bayar", "prm_pay", "desc" };
                    arrHeader = new string[] { "No", "Tertanggung", "UP", "Premi Dasar", "Tgl.Surender", "Nilai Surender", "Bank" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["nam"].ToString();
                        arrData[2] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[3] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[4] = dt.Rows[i]["Tgl_bayar"].ToString();
                        arrData[5] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm_pay"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[6] = dt.Rows[i]["desc"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = true;
                }
                else if (fileName.IndexOf("cimbpls") > -1)
                {
                    arrColumn = new string[] { "no", "npol", "ppol", "premigross", "preminetto", "billno" };
                    arrHeader = new string[] { "No", "Policy No.", "Pemegang Polis", "Premi Gross", "Premi netto", "Desc" };

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = x.ToString();
                        arrData[1] = dt.Rows[i]["npol"].ToString();
                        arrData[2] = dt.Rows[i]["PPOL"].ToString();
                        arrData[3] = "Rp. " + Convert.ToInt64(dt.Rows[i]["premigross"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[4] = "Rp. " + Convert.ToInt64(dt.Rows[i]["preminetto"].ToString()).ToString("N", CultureInfo.GetCultureInfo("id-ID").NumberFormat);
                        arrData[5] = dt.Rows[i]["billno"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = true;
                }

                else if (fileName.IndexOf("srbri") > -1)   //claim
                {
                    arrColumn = new string[] { "npol", "bcode", "npst", "nam", "idat", "sex", "addr_1", "addr_2", "addr_3", "addr_4","city","zip","barcode","tgl_cetak",
                                               "umur","prm","up","dob","endorse","polinduk","npol2"};
                    arrHeader = new string[] { "npol", "bcode", "npst", "nam", "idat", "sex", "addr_1", "addr_2", "addr_3", "addr_4","city","zip","barcode","tgl_cetak",
                                               "umur","prm","up","dob","endorse","polinduk","npol2"};

                    dtCust = PopulateColumn(arrColumn, arrHeader);

                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = dt.Rows[i]["npol"].ToString();
                        arrData[1] = dt.Rows[i]["bcode"].ToString();
                        arrData[2] = Convert.ToInt64(dt.Rows[i]["npst"]).ToString();
                        arrData[3] = dt.Rows[i]["nam"].ToString();
                        arrData[4] = Convert.ToDateTime(dt.Rows[i]["idat"].ToString()).ToString("dd-MM-yyyy");
                        arrData[5] = dt.Rows[i]["sex"].ToString();
                        arrData[6] = dt.Rows[i]["addr_1"].ToString();
                        arrData[7] = dt.Rows[i]["addr_2"].ToString();
                        arrData[8] = dt.Rows[i]["addr_3"].ToString();
                        arrData[9] = dt.Rows[i]["addr_4"].ToString();
                        arrData[10] = dt.Rows[i]["city"].ToString();
                        arrData[11] = Convert.ToInt64(dt.Rows[i]["zip"]).ToString();
                        arrData[12] = dt.Rows[i]["barcode"].ToString();
                        arrData[13] = Convert.ToDateTime(dt.Rows[i]["tgl_cetak"].ToString()).ToString("dd-MM-yyyy");
                        arrData[14] = Convert.ToInt64(dt.Rows[i]["umur"]).ToString();
                        arrData[15] = Convert.ToInt64(dt.Rows[i]["prm"]).ToString();
                        arrData[16] = Convert.ToInt64(dt.Rows[i]["up"]).ToString();
                        arrData[17] = Convert.ToDateTime(dt.Rows[i]["dob"].ToString()).ToString("dd-MM-yyyy");
                        arrData[18] = dt.Rows[i]["endorse"].ToString();
                        arrData[19] = dt.Rows[i]["polinduk"].ToString();
                        arrData[20] = dt.Rows[i]["npol2"].ToString();
                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = true;
                }
                else if (fileName.IndexOf("ajkres") > -1)
                {
                    //arrColumn = new string[] {"npol", "nam", "umur","idat", "up", "prm", "mdatnew", "op"};
                    //arrHeader = new string[] {"npol", "nama","umur","idat", "up", "premi", "mdatnew", "opt"};

                    arrColumn = new string[] {"npol", "nam","umur","up","idat","prm","mdatnew","bank"};
                    arrHeader = new string[] {"npol", "nama","umur","up","idat","premi","New Date","Bank"};

                    dtCust = PopulateColumn(arrColumn, arrHeader);
                    var x = 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] arrData = new string[dtCust.Columns.Count];
                        arrData[0] = dt.Rows[i]["npol"].ToString();
                        arrData[1] = dt.Rows[i]["nam"].ToString();
                        arrData[2] = Convert.ToInt32(dt.Rows[i]["umur"]).ToString();
                        arrData[3] = "Rp. " + Convert.ToInt64(dt.Rows[i]["up"].ToString());
                        arrData[4] = Convert.ToDateTime(dt.Rows[i]["idat"]).ToString("dd/MM/yyyy");
                        arrData[5] = "Rp. " + Convert.ToInt64(dt.Rows[i]["prm"].ToString());
                        arrData[6] = Convert.ToDateTime(dt.Rows[i]["mdatnew"]).ToString("dd/MM/yyyy");
                        arrData[7] = dt.Rows[i]["bank"].ToString();

                        dtCust.Rows.Add(arrData);
                        x++;
                    }

                    chkPDF.Visible = true;
                }


                gvData.DataSource = dtCust;
                gvData.DataBind();

                var asd = gvData.Columns.Count;
                gvData.FooterRow.Cells[asd - 2].Text = "<b>Total Data</b>";
                gvData.FooterRow.Cells[asd - 2].HorizontalAlign = HorizontalAlign.Right;
                gvData.FooterRow.Cells[asd - 1].Text = "<b>" + dtCust.Rows.Count.ToString() + "</b>";

                gvData.Visible = true;
                btnDownload.Visible = true;
                btnCancel.Visible = true;

                if (dtCust.Rows.Count < 1) { btnDownload.Enabled = false; } else { Session["CustData"] = dt; }
            }
            catch (Exception ex)
            {
                Log.Error(ex);

                ShowAlert("Failed. Error when synchronize. " + ex.Message, "warn");
            }
        }
        else
        {
            ShowAlert("Downloaded File Not Found", "warn");
        }
    }

    private string DownloadFile(string ftpSource, string filePath, string user, string pass)
    {
        int intByteRead = 0;
        byte[] byteBuffer = new byte[2048];

        FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(new Uri(ftpSource));
        ftpRequest.Proxy = null;
        ftpRequest.UsePassive = true;
        ftpRequest.UseBinary = true;
        ftpRequest.KeepAlive = true;
        ftpRequest.Credentials = new NetworkCredential(user, pass);
        ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

        var strStatus = "";
        try
        {
            Stream reader = ftpRequest.GetResponse().GetResponseStream();
            FileStream fileStream = new FileStream(filePath, FileMode.Create);
            while (true)
            {
                intByteRead = reader.Read(byteBuffer, 0, byteBuffer.Length);
                if (intByteRead == 0)
                    break;

                fileStream.Write(byteBuffer, 0, intByteRead);
            }

            fileStream.Close();
            strStatus = "Download Success";
        }
        catch (WebException ex)
        {
            strStatus = ((FtpWebResponse)ex.Response).StatusDescription;
        }

        return strStatus;
    }

    private void Clear()
    {
        chkPDF.Visible = false;
        gvData.Visible = false;
        btnCancel.Visible = false;
        btnDownload.Visible = false;
    }

    protected void btnDownload_Click(object sender, EventArgs e)
    {
        try
        {
            var strLetter = hdnLetter.Value;
            var strPDF = chkPDF.Checked ? "PDF" : "SURAT";

            var strPath = Main.ServerPath() + @"TempFile\";

            var strFname = "";
            switch (strLetter)
            {
                case "SA":
                    strFname = "ACC";
                    break;
                case "SE":
                    strFname = "ACC-extra";
                    break;
                case "SS":
                    strFname = "Surender";
                    break;
                case "WS":
                    strFname = "WRKSH";
                    break;
                case "SC":
                    strFname = "claim";
                    break;
                case "SP":
                    strFname = "postpone";
                    break;
                case "SD":
                    strFname = "decline";
                    break;
                case "SM":
                    strFname = "medical";
                    break;
                case "PL":
                    strFname = "cimbpl";
                    break;
                case "LP":                      //List peserta 
                    strFname = "listcr";
                    break;
                case "NS":
                    strFname = "OCBC_NISP_NOTA_KTA";     //***amad***
                    break;
                case "SB":
                    strFname = "bpp";     //***amad***
                    break;
                case "AJ":
                    strFname = "ajk_restrukturisasi";
                    break;
            }

            var strLocRpt = Main.ServerPath() + @"CrystalReport\Letter\" + strFname + ".rpt";

            ReportDocument rd = new ReportDocument();
            rd.Load(strLocRpt);

            DataTable dtCust = (DataTable)Session["CustData"];
            if (dtCust.Rows.Count > 0)
            {
                GeneratePDF(rd, dtCust, strLetter, strFname);
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex);
            ShowAlert(ex.Message + " Please contact Administrator.");
        }
    }

    private void GeneratePDF(ReportDocument rd, DataTable dtCustomer, string type, string letterName)
    {
        var arrLetterHeader = new string[] { "tgl_surat", "no_surat", "nama", "desc", "alamat", "kota", "fax", "zip", "companyid" };
        var arrLetterFooter = new string[] { "company", "pic_name", "pic_level1", "pic_level2" };
        var arrLetterBody = new string[0];
        StringBuilder sb = new StringBuilder(); //@tri2
        var arrCustomer = new string[] { "nam", "dob", "umur", "up", "ctr" };

        var strPath = Main.ServerPath() + @"TempFile\" + "mriletter\\" + Session["userName"].ToString() + "\\";
        if (!Directory.Exists(strPath))
        {
            Directory.CreateDirectory(strPath);
        }
        else
        {
            Array.ForEach(Directory.GetFiles(strPath), File.Delete);
        }
        var strFilename = "";
        var strNpol = "";
        var strNam = "";
        var strDate = "";

        CultureInfo ci = new CultureInfo("id-ID");

        List<ListItem> listFile = new List<ListItem>();

        for (int i = 0; i < dtCustomer.Rows.Count; i++)
        {
            if (type == "LP")  //List peserta 
            {
                arrLetterBody = new string[] 
                {
                    "action","NPOL","BCODE","NPST","NAM","CIF","UMUR","UP","PRM","EPRM","IDAT","MDAT","DESC","PDISC","NETT","NETTO","TGL_SURAT", "bank","no_surat", "nama", "desc", "alamat", "kota", "fax", "zip", "companyid" 
                };
                var arrLetter = new string[arrLetterBody.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);

                //for (int j = 0; j < arrLetter.Length; j++)
                //{
                //    var strParamCol = arrLetter[j];
                //    var strParamVal = new object();

                //    if (strParamCol == "action")
                //    {
                //        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                //    }
                //    else
                //    {
                //        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                //    }


                //    //rd.SetParameterValue(strParamCol, strParamVal);
                //}
                rd.SetDataSource(dtCustomer);
            }
            else if (type == "PL")  //Invoice CIMB-PL
            {
                arrLetterBody = new string[] 
                {
                    "action","Billno", "NPOL","PPOL","BILLNO","ALAMAT1","ALAMAT2","ALAMAT3","ALAMAT4","TELP","BANK_PIC","REK1","REK2","REK3","TGL_SURAT","TTD1","TTD2","TTD3","PERIODE","PREMIGROSS","KOMISI_GRS","KOMISI_NET","KOMISI_PPN","MRKCOS_GRS","MRKCOS_NET","MRKCOS_PPN","PPN_TOTAL","KOMISI_TOT","PREMINETTO"
                };
                var arrLetter = new string[arrLetterBody.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                //Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                //Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                //Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }


                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            else if (type == "NS")  //Invoice KTA OCBC NISP  ***amad***
            {
                arrLetterBody = new string[] { "comp", "alm1", "alm2", "alm3", "alm4", "alm5", "prmgrs", "pdisc", "tprm", "norek", "pic", "bl", "th", "pejabat", "level", "billdate2" };
                var arrLetter = new string[arrLetterBody.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();
                    if (strParamCol == "comp")
                    {
                        strParamVal = dtCustomer.Rows[i]["comp"].ToString();
                    }
                    else if (strParamCol == "alm1")
                    {
                        strParamVal = dtCustomer.Rows[i]["alm1"].ToString();
                    }
                    else if (strParamCol == "alm2")
                    {
                        strParamVal = dtCustomer.Rows[i]["alm2"].ToString();
                    }
                    else if (strParamCol == "alm3")
                    {
                        strParamVal = dtCustomer.Rows[i]["alm3"].ToString();
                    }
                    else if (strParamCol == "alm4")
                    {
                        strParamVal = dtCustomer.Rows[i]["alm4"].ToString();
                    }
                    else if (strParamCol == "alm5")
                    {
                        strParamVal = dtCustomer.Rows[i]["alm5"].ToString();
                    }
                    else if (strParamCol == "prmgrs")
                    {
                        strParamVal = dtCustomer.Rows[i]["prmgrs"];
                    }
                    else if (strParamCol == "pdisc")
                    {
                        strParamVal = dtCustomer.Rows[i]["pdisc"];
                    }
                    else if (strParamCol == "tprm")
                    {
                        strParamVal = dtCustomer.Rows[i]["tprm"];
                    }
                    else if (strParamCol == "norek")
                    {
                        strParamVal = dtCustomer.Rows[i]["norek"].ToString();
                    }
                    else if (strParamCol == "pic")
                    {
                        strParamVal = dtCustomer.Rows[i]["pic"].ToString();
                    }
                    else if (strParamCol == "bl")
                    {
                        strParamVal = dtCustomer.Rows[i]["bl"];
                    }
                    else if (strParamCol == "th")
                    {
                        strParamVal = dtCustomer.Rows[i]["th"];
                    }
                    else if (strParamCol == "pejabat")
                    {
                        strParamVal = dtCustomer.Rows[i]["pejabat"].ToString();
                    }
                    else if (strParamCol == "level")
                    {
                        strParamVal = dtCustomer.Rows[i]["level"].ToString();
                    }
                    else if (strParamCol == "billdate2")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i]["billdate2"].ToString());
                    }
                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            if (type == "SD")  //decline
            {
                arrLetterBody = new string[] 
                {
                    "action","no_surat"
                };

                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "tgl_surat" || strParamCol == "dob")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                    }
                    else if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "desc")
                    {
                        strParamVal = dtCustomer.Rows[i]["desc"].ToString();
                    }
                    else if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            if (type == "SM")  //medical
            {
                arrLetterBody = new string[] 
                {
                    "action","alamat1","no_surat","med_req1","med_req2","med_req3","med_req4"
                };

                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "tgl_surat" || strParamCol == "dob")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                    }
                    else if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "desc")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat1"].ToString() == "" ? dtCustomer.Rows[i]["desc"].ToString() : dtCustomer.Rows[i]["alamat1"].ToString();
                    }
                    else if (strParamCol == "alamat")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat2"].ToString() == "" ? dtCustomer.Rows[i]["alamat"].ToString() : dtCustomer.Rows[i]["alamat2"].ToString();
                    }
                    else if (strParamCol == "alamat1")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat3"].ToString();
                    }
                    else if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            if (type == "SP")  //postpone
            {
                arrLetterBody = new string[] 
                {
                    "action","no_surat","remarks","tindakan"
                };

                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "tgl_surat" || strParamCol == "dob")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                    }
                    else if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "desc")
                    {
                        strParamVal = dtCustomer.Rows[i]["desc"].ToString();
                    }
                    else if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }

            else if (type == "SA")   //akseptasi
            {
                arrLetterBody = new string[] {
                    "action","alamat1","rate","premigross","preminetto","komisinet","acc_name","acc_addr","acc_no","metode"
                };

                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);
                long cicilan = 0;

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "tgl_surat" || strParamCol == "dob")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                    }
                    else if (strParamCol == "up" || strParamCol == "premigross" || strParamCol == "preminetto" || strParamCol == "komisinet" || strParamCol == "preminetto")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "desc")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat1"].ToString() == "" ? dtCustomer.Rows[i]["desc"].ToString() : dtCustomer.Rows[i]["alamat1"].ToString();
                    }
                    else if (strParamCol == "alamat")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat2"].ToString() == "" ? dtCustomer.Rows[i]["alamat"].ToString() : dtCustomer.Rows[i]["alamat2"].ToString();
                    }
                    else if (strParamCol == "alamat1")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat3"].ToString();
                    }
                    else if (strParamCol == "rate")
                    {
                        if (dtCustomer.Rows[i]["instalment"].ToString() != "0")
                        {
                            strParamVal = dtCustomer.Rows[i]["rate"].ToString() + " + charge " + Convert.ToInt64(dtCustomer.Rows[i]["instfactor"].ToString()) + "%";
                        }
                        else
                        {
                            if (dtCustomer.Rows[i]["companyid"].ToString() == "1")  //"{0:N4}"
                            {
                                //strParamVal = string.Format("{0:n0}",dtCustomer.Rows[i]["rate"].ToString());
                               // strParamVal = String.Format("{0:n0}", Convert.ToInt64(dtCustomer.Rows[i]["rate"]));
                                strParamVal = dtCustomer.Rows[i]["rate"].ToString();
                            }
                            else
                            {
                                //strParamVal = string.Format("{0:n4}", dtCustomer.Rows[i]["rate"].ToString());
                               // strParamVal = String.Format("{0:n4}", Convert.ToDouble(dtCustomer.Rows[i]["rate"]));
                                strParamVal = dtCustomer.Rows[i]["rate"].ToString();
                            }

                        }
                    }
                    else if (strParamCol == "metode")
                    {
                        if (dtCustomer.Rows[i]["instalment"].ToString() != "0")
                        {
                            cicilan = Convert.ToInt64(dtCustomer.Rows[i]["preminetto"]) / Convert.ToInt64(dtCustomer.Rows[i]["instalment"]);
                            strParamVal = "Premi Cicilan  Rp." + string.Format("{0:n0}", cicilan) + " (" + Convert.ToInt64(dtCustomer.Rows[i]["instalment"].ToString()) + "x)";
                        }
                        else
                        {
                            strParamVal = "Premi Sekaligus";
                        }
                    }
                    else if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            else if (type == "WS")
            {
                arrLetterBody = new string[] 
                { 
                    "indate","desc","sex","med_req","user","rate","prm","erate","eprm","up_total","tb","bb","prev_spa","npol",
                    "prevspa1","previdat1","prevstat1","prevup1","prevctr1","prevextrt1","prevup1_0",
                    "prevspa2","previdat2","prevstat2","prevup2","prevctr2","prevextrt2","prevup2_0",
                    "prevspa3","previdat3","prevstat3","prevup3","prevctr3","prevextrt3","prevup3_0",
                    "prevspa4","previdat4","prevstat4","prevup4","prevctr4","prevextrt4","prevup4_0","profesi","inhour","medical"
                };

                var strPREV = "N";
                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);

                for (int j = 0; j < arrLetter.Count(); j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "indate" || strParamCol == "dob")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd-MMM-yyyy", ci);
                    }
                    else if (strParamCol == "inhour")
                    {
                        strParamVal = dtCustomer.Rows[i]["rectime"].ToString();
                    }
                    else if (strParamCol == "up" || strParamCol == "prm" || strParamCol == "eprm" || strParamCol == "up_total" || strParamCol == "ctr")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "npol")
                    {
                        strParamVal = dtCustomer.Rows[i]["npol"].ToString() + "-" + dtCustomer.Rows[i]["bcode"].ToString();
                        if ((dtCustomer.Rows[i]["npst"].ToString() == "0") || (dtCustomer.Rows[i]["npst"].ToString() == ""))
                        {
                            strParamVal = strParamVal + " . . . . . . . . / " + dtCustomer.Rows[i]["urut"].ToString();
                        }
                        else
                        {
                            strParamVal = strParamVal + "-" + dtCustomer.Rows[i]["npst"].ToString() + " / " + dtCustomer.Rows[i]["urut"].ToString();
                        }
                    }
                    else if (strParamCol == "prev_spa")
                    {
                        if ((strPREV == "N") &&
                           ((dtCustomer.Rows[i]["prevspa1"].ToString().Trim() != "") ||
                           (dtCustomer.Rows[i]["prevspa2"].ToString().Trim() != "") ||
                           (dtCustomer.Rows[i]["prevspa3"].ToString().Trim() != "") ||
                           (dtCustomer.Rows[i]["prevspa4"].ToString().Trim() != "")))
                        {
                            strParamVal = "Y";
                        }
                        else
                        {
                            strParamVal = "N";
                        }

                    }
                    else if (strParamCol == "sex")
                    {
                        strParamVal = dtCustomer.Rows[i]["sex"].ToString() == "True" ? "Male" : "Female";
                    }
                    else if (strParamCol.Substring(0, strParamCol.Length - 1) == "prevspa")
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString().Trim();
                    }
                    else if (strParamCol.Substring(0, strParamCol.Length - 1) == "previdat")
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString() == "" ? "" : Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd-MM-yyyy", ci);
                    }
                    else if (strParamCol.Substring(0, strParamCol.Length - 1) == "prevstat")
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString().Trim();
                    }
                    else if (strParamCol.Substring(0, strParamCol.Length - 1) == "prevup")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol.Substring(0, strParamCol.Length - 1) == "prevctr")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol.Substring(0, strParamCol.Length - 1) == "prevextrt")
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString() == "" ? "0" : dtCustomer.Rows[i][strParamCol].ToString();
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(arrLetter[j], strParamVal);
                    //Log.Trace("arrLetter[" + j + "]:" + arrLetter[j] + "= " + strParamVal);
                }
            }
            else if (type == "SE")
            {
                arrLetterBody = new string[] 
                {
                    "action","alamat1","rate","premitotal","premidasar","premiextra","ratepremiextra","komisipremi","acc_name","acc_addr","acc_no","metode","remark"
                };

                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "tgl_surat" || strParamCol == "dob")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                    }
                    else if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "premidasar")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["prm"].ToString());
                    }
                    else if (strParamCol == "premiextra")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["eprm"].ToString());
                    }
                    else if (strParamCol == "ratepremiextra")
                    {
                        // strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["erate"].ToString());
                        if (dtCustomer.Rows[i]["companyid"].ToString() == "23")
                        {
                            strParamVal = dtCustomer.Rows[i]["rateextra"].ToString();
                        }
                        else
                        {
                            strParamVal = dtCustomer.Rows[i]["erate"].ToString() == "0" ? dtCustomer.Rows[i]["ext"].ToString() : dtCustomer.Rows[i]["erate"].ToString();
                        }
                    }
                    else if (strParamCol == "premitotal")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["prm"].ToString()) + Convert.ToInt64(dtCustomer.Rows[i]["eprm"].ToString());
                    }
                    else if (strParamCol == "komisipremi")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["komisigros"].ToString());
                    }
                    else if (strParamCol == "desc")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat1"].ToString() == "" ? dtCustomer.Rows[i]["desc"].ToString() : dtCustomer.Rows[i]["alamat1"].ToString();
                    }
                    else if (strParamCol == "alamat")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat2"].ToString() == "" ? dtCustomer.Rows[i]["alamat"].ToString() : dtCustomer.Rows[i]["alamat2"].ToString();
                    }
                    else if (strParamCol == "alamat1")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat3"].ToString();
                    }
                    else if (strParamCol == "metode")
                    {
                        strParamVal = "Premi Sekaligus";
                    }
                    else if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else if (strParamCol == "remark")
                    {
                        strParamVal = dtCustomer.Rows[i]["remark"].ToString();
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            else if (type == "SS")
            {
                arrLetterBody = new string[] 
                {
                    "action","alamat1","premitotal","idat","mdat","Surr_amt","tgl_bayar","acc_name","acc_addr","acc_no","no_surat",
                    "NoPolis","Ratepremi","kom_basic","jnskredit","prm","eprm","tgl_pk","poldur","rumus1" ,"rumus2","oleh","ext"      
                };

                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "tgl_surat" || strParamCol == "dob" || strParamCol == "tgl_bayar")
                    {
                        if (dtCustomer.Rows[i][strParamCol].ToString() == "")
                        {
                            strParamVal = "";
                        }
                        else
                        {
                            strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                        }
                    }
                    else if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "Surr_amt")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["PRM_PAY"].ToString());
                    }
                    else if (strParamCol == "premitotal")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["prm"].ToString()) + Convert.ToInt64(dtCustomer.Rows[i]["eprm"].ToString());
                    }
                    else if (strParamCol == "desc")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat1"].ToString() == "" ? dtCustomer.Rows[i]["desc"].ToString() : dtCustomer.Rows[i]["alamat1"].ToString();
                    }
                    else if (strParamCol == "alamat")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat2"].ToString() == "" ? dtCustomer.Rows[i]["alamat"].ToString() : dtCustomer.Rows[i]["alamat2"].ToString();
                    }
                    else if (strParamCol == "alamat1")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat3"].ToString();
                    }
                    else if (strParamCol == "acc_name")
                    {
                        strParamVal = dtCustomer.Rows[i]["atsnama1"].ToString().TrimEnd() + dtCustomer.Rows[i]["atsnama2"].ToString().TrimEnd();
                    }
                    else if (strParamCol == "acc_addr")
                    {
                        strParamVal = dtCustomer.Rows[i]["nmbank1"].ToString().TrimEnd() + " " + dtCustomer.Rows[i]["nmbank2"].ToString().TrimEnd();
                    }
                    else if (strParamCol == "acc_no")
                    {
                        strParamVal = dtCustomer.Rows[i]["norek"].ToString();
                    }
                    else if (strParamCol == "no_surat")
                    {
                        strParamVal = dtCustomer.Rows[i]["nosurat"].ToString();
                    }
                    else if (strParamCol == "NoPolis")
                    {
                        strParamVal = dtCustomer.Rows[i]["npol"].ToString() + "-" + dtCustomer.Rows[i]["bcode"].ToString() + "-" + dtCustomer.Rows[i]["npst"].ToString();
                    }
                    else if (strParamCol == "poldur")
                    {
                        strParamVal = dtCustomer.Rows[i]["pls_age"].ToString().TrimStart();
                    }
                    else if (strParamCol == "ext")
                    {
                        strParamVal = dtCustomer.Rows[i]["ext"].ToString() == "0" ? "-" : dtCustomer.Rows[i]["ext"].ToString().TrimStart() + " %";
                    }
                    else if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            else if (type == "SC")
            {
                arrLetterBody = new string[] 
                {
                    "action","alamat1","idat","mdat","tgl_bayar","acc_name","acc_addr","acc_no","acc_kota","no_surat","sex",
                    "prm","eprm","death","poldur","poldur2","deathcause","up_dec","oleh","death","deathcase","NoPolis","clmbayar","rateamt","mengetahui","diperiksa"  
                };

                var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
                Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
                Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
                Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();

                    if (strParamCol == "tgl_surat" || strParamCol == "dob" || strParamCol == "tgl_bayar")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                    }
                    else if (strParamCol == "sex")
                    {
                        strParamVal = dtCustomer.Rows[i]["sex"].ToString() == "True" ? "Pria" : "Wanita";
                    }
                    else if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i][strParamCol].ToString());
                    }
                    else if (strParamCol == "desc")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat1"].ToString() == "" ? dtCustomer.Rows[i]["desc"].ToString() : dtCustomer.Rows[i]["alamat1"].ToString();
                    }
                    else if (strParamCol == "alamat")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat2"].ToString() == "" ? dtCustomer.Rows[i]["alamat"].ToString() : dtCustomer.Rows[i]["alamat2"].ToString();
                    }
                    else if (strParamCol == "alamat1")
                    {
                        strParamVal = dtCustomer.Rows[i]["alamat3"].ToString();
                    }
                    else if (strParamCol == "acc_name")
                    {
                        strParamVal = dtCustomer.Rows[i]["atsnama1"].ToString().TrimEnd() + dtCustomer.Rows[i]["atsnama2"].ToString().TrimEnd();
                    }
                    else if (strParamCol == "acc_addr")
                    {
                        strParamVal = dtCustomer.Rows[i]["nmbank1"].ToString().TrimEnd();
                    }
                    else if (strParamCol == "acc_kota")
                    {
                        strParamVal = dtCustomer.Rows[i]["nmbank2"].ToString().TrimEnd();
                    }
                    else if (strParamCol == "acc_no")
                    {
                        strParamVal = dtCustomer.Rows[i]["norek"].ToString();
                    }
                    else if (strParamCol == "no_surat")
                    {
                        strParamVal = dtCustomer.Rows[i]["nosurat"].ToString();
                    }
                    else if (strParamCol == "poldur")
                    {
                        strParamVal = dtCustomer.Rows[i]["pls_age"].ToString().TrimStart();
                    }
                    else if (strParamCol == "poldur2")
                    {
                        if (dtCustomer.Rows[i]["deathcause"].ToString() == "")
                        {
                            strParamVal = "";
                        }
                        else
                        {
                            strParamVal = dtCustomer.Rows[i]["deathcause"].ToString().Substring(0, 5) + " bulan dari " + dtCustomer.Rows[i]["deathcause"].ToString().Substring(6, 3);
                        }
                    }
                    else if (strParamCol == "NoPolis")
                    {
                        strParamVal = dtCustomer.Rows[i]["npol"].ToString() + "-" + dtCustomer.Rows[i]["bcode"].ToString() + "-" + dtCustomer.Rows[i]["npst"].ToString();
                    }
                    else if (strParamCol == "clmbayar")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["clm_actual"].ToString());
                    }
                    else if (strParamCol == "rateamt")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["rateamt"].ToString());
                    }
                    else if (strParamCol == "mengetahui")
                    {
                        strParamVal = dtCustomer.Rows[i]["mengetahui"].ToString();
                    }
                    else if (strParamCol == "diperiksa")
                    {
                        strParamVal = dtCustomer.Rows[i]["diperiksa"].ToString();
                    }
                    else if (strParamCol == "action")
                    {
                        strParamVal = chkPDF.Checked ? "PDF" : "SURAT";
                    }
                    else
                    {
                        strParamVal = dtCustomer.Rows[i][strParamCol].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                }
            }
            else if (type == "SB")
            {
                arrLetterBody = new string[] 
                {
                    "npol", "bcode", "npst", "nam", "addr_1", "addr_2","addr_3","addr_4","city","zip","barcode","tgl_cetak","idat","sex",
                    "umur","prm","up","dob","endorse","polinduk","npol2"
                };

              //  var arrLetter = new string[arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length + arrLetterFooter.Length];
                var arrLetter = new string[arrLetterBody.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);
           //     Array.Copy(arrCustomer, 0, arrLetter, arrLetterBody.Length, arrCustomer.Length);
           //     Array.Copy(arrLetterHeader, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length, arrLetterHeader.Length);
           //     Array.Copy(arrLetterFooter, 0, arrLetter, arrLetterBody.Length + arrCustomer.Length + arrLetterHeader.Length, arrLetterFooter.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();
                    if (strParamCol == "npol")
                    {
                        strParamVal = dtCustomer.Rows[i]["npol"].ToString();
                    }
                    if (strParamCol == "bcode")
                    {
                        strParamVal = dtCustomer.Rows[i]["bcode"].ToString();
                    }
                    if (strParamCol == "npst")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["npst"].ToString());
                    }
                    if (strParamCol == "umur")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["umur"].ToString());
                    }
                    if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["up"].ToString());
                    }
                    if (strParamCol == "prm")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["prm"].ToString());
                    }

                    if (strParamCol == "tgl_cetak" || strParamCol == "idat" || strParamCol == "dob")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i][strParamCol].ToString()).ToString("dd MMMM yyyy", ci);
                    }
                    else if (strParamCol == "sex")
                    {
                        strParamVal = dtCustomer.Rows[i]["sex"].ToString() == "True" ? "Bapak" : "Ibu";
                    }
                    else if (strParamCol == "nam")
                    {
                        strParamVal = dtCustomer.Rows[i]["nam"].ToString();
                    }
                    else if (strParamCol == "endorse")
                    {
                        strParamVal = dtCustomer.Rows[i]["endorse"].ToString();
                    }
                    else if (strParamCol == "polinduk")
                    {
                        strParamVal = dtCustomer.Rows[i]["polinduk"].ToString();
                    }
                    else if (strParamCol == "npol2")
                    {
                        strParamVal = dtCustomer.Rows[i]["npol2"].ToString();
                    }
                    else if (strParamCol == "addr_1")
                    {
                        strParamVal = dtCustomer.Rows[i]["addr_1"].ToString();
                    }
                    else if (strParamCol == "addr_2")
                    {
                        strParamVal = dtCustomer.Rows[i]["addr_2"].ToString();
                    }
                    else if (strParamCol == "addr_3")
                    {
                        strParamVal = dtCustomer.Rows[i]["addr_3"].ToString();
                    }
                    else if (strParamCol == "addr_4")
                    {
                        strParamVal = dtCustomer.Rows[i]["addr_4"].ToString();
                    }
                    else if (strParamCol == "city")
                    {
                        strParamVal = dtCustomer.Rows[i]["city"].ToString();
                    }
                    else if (strParamCol == "zip")
                    {
                        strParamVal = Convert.ToInt64(dtCustomer.Rows[i]["zip"].ToString());
                    }
                    else if (strParamCol == "barcode")
                    {
                        strParamVal = dtCustomer.Rows[i]["barcode"].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                    Log.Trace(strParamCol + " " + strParamVal);
                }
            }
            else if (type == "AJ")
            {
                arrLetterBody = new string[] { "npol", "nam","umur","up","idat","prm","mdatnew","op","bank"};

                var arrLetter = new string[arrLetterBody.Length];
                Array.Copy(arrLetterBody, arrLetter, arrLetterBody.Length);

                for (int j = 0; j < arrLetter.Length; j++)
                {
                    var strParamCol = arrLetter[j];
                    var strParamVal = new object();
                    if (strParamCol == "npol")
                    {
                        strParamVal = dtCustomer.Rows[i]["npol"].ToString();
                    }

                    if (strParamCol == "nam")
                    {
                        strParamVal = dtCustomer.Rows[i]["nam"].ToString();
                    }

                    if (strParamCol == "umur")
                    {
                        strParamVal = Convert.ToInt32(dtCustomer.Rows[i]["umur"]).ToString();
                    }

                    if (strParamCol == "up")
                    {
                        strParamVal = Convert.ToInt32(dtCustomer.Rows[i]["up"]).ToString();
                    }

                    if (strParamCol == "idat")
                    {
                       strParamVal = Convert.ToDateTime(dtCustomer.Rows[i]["idat"]).ToString("dd/MM/yyyy");
                    }

                    if (strParamCol == "prm")
                    {
                        strParamVal = Convert.ToInt32(dtCustomer.Rows[i]["prm"]).ToString();
                    }

                    if (strParamCol == "mdatnew")
                    {
                        strParamVal = Convert.ToDateTime(dtCustomer.Rows[i]["mdatnew"]).ToString("dd/MM/yyyy");
                    }

                    if (strParamCol == "op")
                    {
                        strParamVal = Convert.ToInt16(dtCustomer.Rows[i]["op"]).ToString();
                    }

                    if (strParamCol == "bank")
                    {
                        strParamVal = dtCustomer.Rows[i]["bank"].ToString();
                    }

                    rd.SetParameterValue(strParamCol, strParamVal);
                } 

 
            }



            if ((type == "SA") || (type == "SE") || (type == "SM") || (type == "SP") || (type == "SD"))  
            {
                strNpol = dtCustomer.Rows[i]["npol"].ToString();
                strNpol = strNpol + "-" + dtCustomer.Rows[i]["bcode"].ToString();
                strNpol = strNpol + "-" + dtCustomer.Rows[i]["urut"].ToString();
                strNam = dtCustomer.Rows[i]["nam"].ToString();
            }
            //else if ((type == "PL"))
            //{
            //    strNpol = "Billing_CIMB-PL_" + dtCustomer.Rows[i]["periode"].ToString();
            //    strNam = "";
            //}
            else if ((type == "PL") || (type == "NS") || (type == "AJ"))             // ***amad***
            {
                if (type == "PL")
                {
                    strNpol = "Billing_CIMB-PL_" + dtCustomer.Rows[i]["periode"].ToString();
                    strNam = "";
                }
                else if (type == "NS")         // ***amad***
                {
                    strNpol = "Invoice_OCBC_KTA" + dtCustomer.Rows[i]["th"].ToString() + dtCustomer.Rows[i]["bl"].ToString();
                    strNam = "";
                }
                else
                {
                    strNpol = "AJK" + dtCustomer.Rows[i]["npol"].ToString()+"-2020";
                    strNam = ""; 
                }

            }

            else
            {
                strNpol = dtCustomer.Rows[i]["npol"].ToString();
                strNpol = strNpol + "-" + dtCustomer.Rows[i]["bcode"].ToString();
                strNpol = strNpol + "-" + dtCustomer.Rows[i]["npst"].ToString();
                strNam = dtCustomer.Rows[i]["nam"].ToString();
            }

            if (type == "WS")
            {
                strDate = Convert.ToDateTime(dtCustomer.Rows[i]["tgl_stat"].ToString()).ToString("ddMMyyyy");
                strFilename = "WRKSH_" + strNpol + "_" + strNam.Replace(" ", "_") + "_" + strDate + ".pdf";
            }

            else
            {
                if ((type == "PL") || (type == "NS") || (type == "AJ"))         // ***amad***
                {
                    if (type == "PL" || type == "AJ")
                    {
                        strDate = DateTime.Now.ToString("ddMMyyyy");
                        strFilename = strNpol + "_" + strNam.Replace(" ", "_") + "_" + strDate + ".pdf";
                    }
                    else
                    {
                        strFilename = strNpol + ".pdf";           // ***amad***
                    }
                }
                else if ((type == "LP"))         //list Peserta
                {
                    strFilename = "LP_" + dtCustomer.Rows[i]["bank"].ToString() + ".pdf";
                }
                else
                {
                    if ((type == "SB"))
                    {
                        strDate = Convert.ToDateTime(dtCustomer.Rows[i]["tgl_cetak"].ToString()).ToString("ddMMyyyy");
                    }
                    else
                    {
                        strDate = Convert.ToDateTime(dtCustomer.Rows[i]["tgl_surat"].ToString()).ToString("ddMMyyyy");
                    }

                    strFilename = strNpol + "_" + strNam.Replace(" ", "_") + "_" + strDate + ".pdf";
                }
            }
            strFilename = RemoveSpecialCharacters(strFilename);
            if ((strFilename.Substring(0,6)=="201905"||strFilename.Substring(0,6)=="201503") && (type=="SA"))
            {
                int lenfile = strFilename.Trim().Length;
                rd.ExportToDisk(ExportFormatType.PortableDocFormat, strPath + "\\" + strFilename);
//              Response.AddHeader("REFRESH", "2;URL=split_file.aspx");
//              Split_file(Main.ServerPath() + @"TempFile\" + "mriletter\\" + Session["userName"].ToString(), strFilename);
                DownloadFile(Main.ServerPath() + @"TempFile\" + "mriletter\\" + Session["userName"].ToString(), strFilename);

            }
            else
            {
               ConvertStream(rd, strPath + strFilename);
            }

            rd.Refresh();
            if (chkZIP.Checked)
            {
                listFile.Add(new ListItem(strPath + strFilename));
            }
            else
            {
                if (strFilename.Substring(0, 6) != "201905" && strFilename.Substring(0, 6) != "201503")
                {
                    sb.AppendLine("window.open('../TempFile/" + "mriletter/" + Session["userName"].ToString() + "/" + strFilename + "', '_blank');"); //new tab //@tri2
                }
            }
        }

      rd.Close();
    //rd.Dispose();


        if ((listFile.Count > 0) && (chkZIP.Checked))
        {
            var zipName = String.Format(letterName + "_Letter_{0}.zip", DateTime.Now.ToString("ddMMyyyy_hhmmssfff"));
            DownloadAsZip(listFile, zipName);
        }
        else
        {
            this.ClientScript.RegisterStartupScript(this.GetType(), "OpenWindow", sb.ToString(), true); //new tab //@tri2
        }

    }

    private void DownloadFile(string FilePath, string filenamezip)
    {
        var encpathfile = Main.Encryption.Encrypt(FilePath);
        Response.AddHeader("REFRESH", "2;URL=Expense_DownloadFile.aspx?id=" + Server.UrlEncode(encpathfile) + "&xm=HPAllowance");
    }

    private void Split_file(string FilePath, string filename)
    {
        var encpathfile = Main.Encryption.Encrypt(FilePath);
        Response.AddHeader("REFRESH", "2;URL=split_file.aspx?id=" + Server.UrlEncode(encpathfile) + "&xm=HPAllowance");
    }

    private void ConvertStream(ReportDocument rd, string pathFile)
    {
        var streamPDF = rd.ExportToStream(ExportFormatType.PortableDocFormat);

        using (var fileStream = new FileStream(pathFile, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            PdfReader pdfReader = new PdfReader(streamPDF);
            PdfEncryptor.Encrypt(pdfReader, fileStream, true, null, "secret", PdfWriter.AllowPrinting | PdfWriter.AllowScreenReaders);
        }

    }

    private void DownloadAsZip(List<ListItem> listFile, string fileName)
    {
        using (ZipFile zip = new ZipFile())
        {
            zip.AlternateEncodingUsage = ZipOption.AsNecessary;

            for (int i = 0; i < listFile.Count; i++)
            {
                zip.AddFile(listFile[i].Text, "");
            }

            Response.Clear();
            Response.BufferOutput = false;
            Response.ContentType = "application/zip";
            Response.AddHeader("content-disposition", "attachment; filename=" + fileName);
            zip.Save(Response.OutputStream);

            for (int i = 0; i < listFile.Count; i++)
            {
                //File.Delete(listFile[i].Value + listFile[i].Text);
                File.Delete(listFile[i].Value );
            }

            Response.Flush();
            Response.SuppressContent = true;
            ApplicationInstance.CompleteRequest();
        }
    }

    private DataTable PopulateColumn(string[] column, string[] header)
    {
        DataTable dtCust = new DataTable();

        var x = 0;
        foreach (var strColumn in column)
        {
            dtCust.Columns.Add(strColumn);

            BoundField bf = new BoundField();
            bf.DataField = strColumn;
            bf.HeaderText = header[x];
            gvData.Columns.Add(bf);

            x++;
        }

        return dtCust;
    }

    protected void btnCancel_Click(object sender, EventArgs e)
    {
        Clear();
    }

    /// <summary>
    /// Showing alert (Info/Succes/Warning)
    /// </summary>
    /// <param name="strMsg">Your Message</param>
    /// <param name="strType">info/success/warn (Default : info)</param>
    private void ShowAlert(string strMsg, string strType = "info")
    {
        Site master = (Site)this.Master;
        master.Alert(strMsg, strType);
    }

    /// <summary>
    /// Inserting log activity
    /// </summary>
    /// <param name="id">user id from session user_id</param>
    /// <param name="act">activity</param>
    private void AddLog(string id, string act)
    {
        var strPage = HttpContext.Current.Request.Url.AbsolutePath;
        var strIp = Request.UserHostAddress;
        Log.Activity(id, act, strPage, strIp);
    }

    public static string RemoveSpecialCharacters(string str)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < str.Length; i++)
        {
            if ((str[i] >= '0' && str[i] <= '9')
                || (str[i] >= 'A' && str[i] <= 'z'
                    || (str[i] == '.' || str[i] == '_' || str[i] == '-' || str[i] == ',')))
            {
                sb.Append(str[i]);
            }
        }
        return sb.ToString();
    }
}