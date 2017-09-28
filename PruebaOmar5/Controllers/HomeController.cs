using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using System.Data.SqlClient;
using System.Configuration;

namespace PruebaOmar5.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            try
            {
                DataSet ds = new DataSet();
                if (Request.Files["file"].ContentLength > 0)
                {
                    string fileExtension =
                                         System.IO.Path.GetExtension(Request.Files["file"].FileName);

                    if (fileExtension == ".xls" || fileExtension == ".xlsx")
                    {
                        string fileLocation = Server.MapPath("~/Content/") + Request.Files["file"].FileName;
                        if (System.IO.File.Exists(fileLocation))
                        {

                            System.IO.File.Delete(fileLocation);
                        }
                        Request.Files["file"].SaveAs(fileLocation);
                        string excelConnectionString = string.Empty;
                        excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        //connection String for xls file format.
                        if (fileExtension == ".xls")
                        {
                            excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        }
                        //connection String for xlsx file format.
                        else if (fileExtension == ".xlsx")
                        {

                            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        }
                        //Create Connection to Excel work book and add oledb namespace
                        OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                        excelConnection.Open();
                        DataTable dt = new DataTable();

                        dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        if (dt == null)
                        {
                            return null;
                        }

                        String[] excelSheets = new String[dt.Rows.Count];
                        int t = 0;
                        //excel data saves in temp file here.
                        foreach (DataRow row in dt.Rows)
                        {
                            excelSheets[t] = row["TABLE_NAME"].ToString();
                            t++;
                        }
                        OleDbConnection excelConnection1 = new OleDbConnection(excelConnectionString);


                        string query = string.Format("Select * from [{0}]", excelSheets[0]);
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection1))
                        {
                            dataAdapter.Fill(ds);
                        }
                    }
                    if (fileExtension.ToString().ToLower().Equals(".xml"))
                    {
                        string fileLocation = Server.MapPath("~/Content/") + Request.Files["FileUpload"].FileName;
                        if (System.IO.File.Exists(fileLocation))
                        {
                            System.IO.File.Delete(fileLocation);
                        }

                        Request.Files["FileUpload"].SaveAs(fileLocation);
                        XmlTextReader xmlreader = new XmlTextReader(fileLocation);
                        // DataSet ds = new DataSet();
                        ds.ReadXml(xmlreader);
                        xmlreader.Close();
                    }

                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        string conn = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
                        SqlConnection con = new SqlConnection(conn);
                        string query = "INSERT INTO DataPpto (ano_eje, nivel_gob, sector, pliego, u_ejecutora, sec_ejec, programa_pptal, tipo_prod_proy, producto_proyecto, tipo_act_obra_ac, activ_obra_accinv, funcion, division_fn, grupo_fn, meta, finalidad, unidad_medida, cant_meta_anual, cant_meta_sem, avan_fisico_anual, avan_fisico_sem, sec_func, departamento_meta, provincia_meta, distrito_meta, fuente_financ, rubro, categoria_gasto, tipo_transaccion, generica, subgenerica, subgenerica_det, especifica, especifica_det, tipo_recurso, mto_pia, mto_modificaciones, mto_pim, mto_certificado, mto_compro_anual, mto_at_comp_01, mto_at_comp_02, mto_at_comp_03, mto_at_comp_04, mto_at_comp_05, mto_at_comp_06, mto_at_comp_07, mto_at_comp_08, mto_at_comp_09, mto_at_comp_10, mto_at_comp_11, mto_at_comp_12, mto_devenga_01, mto_devenga_02, mto_devenga_03, mto_devenga_04, mto_devenga_05, mto_devenga_06, mto_devenga_07, mto_devenga_08, mto_devenga_09, mto_devenga_10, mto_devenga_11, mto_devenga_12, mto_girado_01, mto_girado_02, mto_girado_03, mto_girado_04, mto_girado_05, mto_girado_06, mto_girado_07, mto_girado_08, mto_girado_09, mto_girado_10, mto_girado_11, mto_girado_12, mto_pagado_01, mto_pagado_02, mto_pagado_03, mto_pagado_04, mto_pagado_05, mto_pagado_06, mto_pagado_07, mto_pagado_08, mto_pagado_09, mto_pagado_10, mto_pagado_11, mto_pagado_12) VALUES ('" + ds.Tables[0].Rows[i][0].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][1].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][2].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][3].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][4].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][5].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][6].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][7].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][8].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][9].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][10].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][11].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][12].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][13].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][14].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][15].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][16].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][17].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][18].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][19].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][20].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][21].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][22].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][23].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][24].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][25].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][26].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][27].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][28].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][29].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][30].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][31].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][32].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][33].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][34].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][35].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][36].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][37].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][38].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][39].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][40].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][41].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][42].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][43].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][44].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][45].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][46].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][47].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][48].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][49].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][50].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][51].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][52].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][53].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][54].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][55].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][56].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][57].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][58].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][59].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][60].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][61].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][62].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][63].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][64].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][65].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][66].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][67].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][68].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][69].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][70].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][71].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][72].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][73].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][74].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][75].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][76].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][77].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][78].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][79].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][80].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][81].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][82].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][83].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][84].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][85].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][86].ToString().Replace("'", "''") + "','" + ds.Tables[0].Rows[i][87].ToString().Replace("'", "''") + "')";
                        con.Open();
                        SqlCommand cmd = new SqlCommand(query, con);
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }

                    ViewBag.Exito = true;
                }
            }

            catch
            {
                ViewBag.Exito = false;
            }

            return View();
        }

        public ActionResult Exito()
        {
            string conn = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
            SqlConnection con = new SqlConnection(conn);
            string query = "select * from CompDevPorOrganizacion";
            con.Open();
            SqlCommand cmd = new SqlCommand(query, con);
            cmd.ExecuteNonQuery();
            con.Close();

            return View();
        }

        public ActionResult ErrorRegistro()
        {
            return View();
        }

    }
}