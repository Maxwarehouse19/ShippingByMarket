﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Configuration;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
//using System.Data;
using ClosedXML.Excel;
using System.Globalization;

namespace ShippingByMarket
{
    public partial class Form1 : Form
    {
        // variables para uso del programa
        // -------------------------------
        string RutaArchivoGeneracion            ="";
        string RutaArchivoVentasNoFacturado    ="";
        string SalesOrderNumber                 ="";
        string HoldCode                         ="";
        string TotalSales                       ="";
        string SalesSku                         ="";
        string SalesCategoryAtTimeOfSale        ="";
        string UomCode                          ="";
        string UomQuantity                      ="";
        string SalesStatus                      ="";
        string SalesOrderDate                   ="";
        string SalesChannelName                 ="";
        string CustomerName                     ="";
        string FulfillmentSku                   ="";
        string FulfillmentChannelName           ="";
        string FulfillmentChannelType           ="";
        string LinkedFulfillmentChannelName     ="";
        string FulfillmentLocationName          ="";
        string FulfillmentOrderNumber           ="";
        string Quantity                         ="";
        string Sku                              ="";
        string Title                            ="";
        string TotalCost                        ="";
        string Commission                       ="";
        string InventoryCost                    ="";
        string UnitCost                         ="";
        string ServiceCost                      ="";
        string EstimatedShippingCost            ="";
        string ShippingCost                     ="";
        string ShippingPrice                    ="";
        string OverheadCost                     ="";
        string PackageCost                      ="";
        string ProfitLoss                       ="";
        string Carrier                          ="";
        string ShippingServiceLevel             ="";
        string ShippedByUser                    ="";
        string ShippingWeight                   ="";
        string Length                           ="";
        string varWidth                         ="";
        string varHeight                        ="";
        string Weight                           ="";
        string StateRegion                      ="";
        string TrackingNum                      ="";
        string MfrName                          ="";
        string PricingRule                      = "";
        string ActualShippingCost               = "";
        string ActualShipping                   = "";
        string ShippingCostDifference           = "";
        int counter = 0;
        string line;
        int cantidad = 0;
        bool Encontro = false;
        string PalabraCompleta = "";
        int ContadorProgreso = 0;
        string ArchivoLog = "";
        string ReporteLog = "";
        
        bool FlgSihayFedex = false;
        bool FlgSihayUSPS = false;
        bool FlgSihayUPS = false;
        bool FlgSihayMI15 = false;
        bool FlgSihayAmazon = false;
        bool EncontroRegistro = false;
        bool FlgSihayEndicia = false;
        string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
        string pathString = "";
        System.IO.StreamReader fileRead = null;

        int contador = 1;
        string BodyExcel = "<html>";

        List<PedidoFedex> listaPedido = new List<PedidoFedex>();

        List<PedidoUSPS> listaPedidoUSPS = new List<PedidoUSPS>();

        List<PedidoUPS> listaPedidoUPS = new List<PedidoUPS>();

        public Form1()
        {
            try
            {
                InitializeComponent();
                //EjecutaProceso();
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error: " + exp.Message);
            }
        }

        private static string GetConnectionString(string file,string Tipo)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension.ToUpper() == "XLS"  )
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";

                if (Tipo== "MASTER")
                    props["Extended Properties"] = "Excel 8.0";
                else
                    props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension.ToUpper() == "XLSX")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";

                if (Tipo == "MASTER")
                    props["Extended Properties"] = "Excel 12.0 XML";
                else
                    props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static DataSet GetDataSetFromExcelFile(string file,string connectionString)
        {
            DataSet ds = new DataSet();

            //string connectionString = GetConnectionString(file,);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        private static DataSet GetDataSetFromExcelFileDetalle(string file, string connectionString)
        {
            DataSet ds = new DataSet();

            //string connectionString = GetConnectionString(file,);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.Contains(" "))
                    {

                        if (!sheetName.EndsWith("$"))
                            continue;
                    }
                    else {
                        if (sheetName.Contains("FilterDatabase"))
                            continue;
                    }

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }


        //  inserta fila a reporte
        // -----------------------
        private void InsertaEncabezadoReporte()
        {
            BodyExcel = "<html>";

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            RutaArchivoGeneracion = pathString +@"\"  ;//ConfigurationManager.AppSettings["RutaArchivoGeneracion"];
            RutaArchivoGeneracion = RutaArchivoGeneracion + "ReporteOutput" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".xls";
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                BodyExcel = BodyExcel + "<body>";
                BodyExcel = BodyExcel + "<table>";
                BodyExcel = BodyExcel + @"<tr bgcolor= ""#CA2229"" style=""color:#ffffff"">";
                // Archivo DW
                // ----------
                BodyExcel = BodyExcel + "<td> SalesOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> TotalSales</td>";
                BodyExcel = BodyExcel + "<td> HoldCode</td>";
                BodyExcel = BodyExcel + "<td> SalesSku</td>";
                BodyExcel = BodyExcel + "<td> SalesCategoryAtTimeOfSale</td>";
                BodyExcel = BodyExcel + "<td> UomCode</td>";
                BodyExcel = BodyExcel + "<td> UomQuantity</td>";
                BodyExcel = BodyExcel + "<td> SalesStatus</td>";
                BodyExcel = BodyExcel + "<td> SalesOrderDate</td>";
                BodyExcel = BodyExcel + "<td> SalesChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentSku</td>";
                BodyExcel = BodyExcel + "<td> CustomerName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> LinkedFulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentLocationName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelType</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> Quantity</td>";
                BodyExcel = BodyExcel + "<td> Sku</td>";
                BodyExcel = BodyExcel + "<td> Title</td>";
                BodyExcel = BodyExcel + "<td> TotalCost</td>";
                BodyExcel = BodyExcel + "<td> Commission</td>";
                BodyExcel = BodyExcel + "<td> InventoryCost</td>";
                BodyExcel = BodyExcel + "<td> UnitCost</td>";
                BodyExcel = BodyExcel + "<td> ServiceCost</td>";
                BodyExcel = BodyExcel + "<td> EstimatedShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingPrice</td>";
                BodyExcel = BodyExcel + "<td> OverheadCost</td>";
                BodyExcel = BodyExcel + "<td> PackageCost</td>";
                BodyExcel = BodyExcel + "<td> ProfitLoss</td>";
                BodyExcel = BodyExcel + "<td> Carrier</td>";
                BodyExcel = BodyExcel + "<td> ShippingServiceLevel</td>";
                BodyExcel = BodyExcel + "<td> ShippedByUser</td>";
                BodyExcel = BodyExcel + "<td> ShippingWeight</td>";
                BodyExcel = BodyExcel + "<td> Weight</td>";
                BodyExcel = BodyExcel + "<td> Width</td>";
                BodyExcel = BodyExcel + "<td> Length</td>";
                BodyExcel = BodyExcel + "<td> Height</td>";
                BodyExcel = BodyExcel + "<td> StateRegion</td>";
                BodyExcel = BodyExcel + "<td> TrackingNum</td>";
                BodyExcel = BodyExcel + "<td> MfrName</td>";
                BodyExcel = BodyExcel + "<td> PricingRule</td>";


                // Archivo Fedex
                // -------------
                //BodyExcel = BodyExcel + "<td>FullTrakingId</td>";
                BodyExcel = BodyExcel + "<td> Ground Tracking ID Prefix</td>";
                BodyExcel = BodyExcel + "<td> Express or Ground Tracking ID</td>";
                BodyExcel = BodyExcel + "<td> Net Charge Amount</td>";
                BodyExcel = BodyExcel + "<td> Service Type</td>";
                BodyExcel = BodyExcel + "<td> Ground Service</td>";
                BodyExcel = BodyExcel + "<td> Shipment Date</td>";
                BodyExcel = BodyExcel + "<td> POD Delivery Date</td>";
                BodyExcel = BodyExcel + "<td> Actual Weight Amount</td>";
                BodyExcel = BodyExcel + "<td> Rated Weight Amount</td>";
                BodyExcel = BodyExcel + "<td> Dim Length</td>";
                BodyExcel = BodyExcel + "<td> Dim Width</td>";
                BodyExcel = BodyExcel + "<td> Dim Height</td>";
                BodyExcel = BodyExcel + "<td> Dim Divisor</td>";
                BodyExcel = BodyExcel + "<td> Shipper State</td>";
                BodyExcel = BodyExcel + "<td> Zone Code</td>";
                BodyExcel = BodyExcel + "<td> Tendered Date</td>";

                // cargos fijos
                // ------------
                BodyExcel = BodyExcel + "<td>Earned Discount</td>";
                BodyExcel = BodyExcel + "<td>Fuel Surcharge</td>";
                BodyExcel = BodyExcel + "<td>Performance Pricing</td>";
                BodyExcel = BodyExcel + "<td>Delivery Area Surcharge Extended</td>";
                BodyExcel = BodyExcel + "<td>Delivery Area Surcharge</td>";
                BodyExcel = BodyExcel + "<td>USPS Non-Mach Surcharge</td>";
                BodyExcel = BodyExcel + "<td>Residential</td>";
                BodyExcel = BodyExcel + "<td>Grace Discount</td>";
                BodyExcel = BodyExcel + "<td>Declared Value</td>";
                BodyExcel = BodyExcel + "<td>DAS Extended Resi</td>";
                BodyExcel = BodyExcel + "<td>Additional Handling</td>";
                BodyExcel = BodyExcel + "<td>Parcel Re-Label Charge</td>";
                BodyExcel = BodyExcel + "<td>Indirect Signature</td>";
                BodyExcel = BodyExcel + "<td>DAS Resi</td>";
                BodyExcel = BodyExcel + "<td>Address Correction</td>";
                BodyExcel = BodyExcel + "<td>DAS Extended Comm</td>";
                BodyExcel = BodyExcel + "<td>Oversize Charge</td>";
                BodyExcel = BodyExcel + "<td>AHS - Dimensions</td>";

                // dato USPS
                BodyExcel = BodyExcel + "<td>Mail Class </td>";
                BodyExcel = BodyExcel + "<td>Tracking Number </td>";
                BodyExcel = BodyExcel + "<td>Total Postage Amt </td>";
                BodyExcel = BodyExcel + "<td>Delivery Date </td>";
                BodyExcel = BodyExcel + "<td>Weight </td>";
                BodyExcel = BodyExcel + "<td>Zone </td>";

                // dato UPS
                BodyExcel = BodyExcel + "<td> Service Type </td>";
                BodyExcel = BodyExcel + "<td>Tracking Number </td>";
                BodyExcel = BodyExcel + "<td>Net Charge Amount </td>";


 
                BodyExcel = BodyExcel + "</tr>";

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";
            }

        }


        //  inserta fila a reporte
        // -----------------------
        private void InsertaEncabezadoVentasNoFacturado()
        {
            BodyExcel = "<html>";

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            RutaArchivoVentasNoFacturado = pathString + @"\";//ConfigurationManager.AppSettings["RutaArchivoGeneracion"];
            RutaArchivoVentasNoFacturado = RutaArchivoVentasNoFacturado + "OutputVentasNoFacturado" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".xls";
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoVentasNoFacturado, true))
            {

                BodyExcel = BodyExcel + "<body>";
                BodyExcel = BodyExcel + "<table>";
                BodyExcel = BodyExcel + @"<tr bgcolor= ""#CA2229"" style=""color:#ffffff"">";
                // Archivo DW
                // ----------
                BodyExcel = BodyExcel + "<td> SalesOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> TotalSales</td>";
                BodyExcel = BodyExcel + "<td> HoldCode</td>";
                BodyExcel = BodyExcel + "<td> SalesSku</td>";
                BodyExcel = BodyExcel + "<td> SalesCategoryAtTimeOfSale</td>";
                BodyExcel = BodyExcel + "<td> UomCode</td>";
                BodyExcel = BodyExcel + "<td> UomQuantity</td>";
                BodyExcel = BodyExcel + "<td> SalesStatus</td>";
                BodyExcel = BodyExcel + "<td> SalesOrderDate</td>";
                BodyExcel = BodyExcel + "<td> SalesChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentSku</td>";
                BodyExcel = BodyExcel + "<td> CustomerName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> LinkedFulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentLocationName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelType</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> Quantity</td>";
                BodyExcel = BodyExcel + "<td> Sku</td>";
                BodyExcel = BodyExcel + "<td> Title</td>";
                BodyExcel = BodyExcel + "<td> TotalCost</td>";
                BodyExcel = BodyExcel + "<td> Commission</td>";
                BodyExcel = BodyExcel + "<td> InventoryCost</td>";
                BodyExcel = BodyExcel + "<td> UnitCost</td>";
                BodyExcel = BodyExcel + "<td> ServiceCost</td>";
                BodyExcel = BodyExcel + "<td> EstimatedShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingPrice</td>";
                BodyExcel = BodyExcel + "<td> OverheadCost</td>";
                BodyExcel = BodyExcel + "<td> PackageCost</td>";
                BodyExcel = BodyExcel + "<td> ProfitLoss</td>";
                BodyExcel = BodyExcel + "<td> Carrier</td>";
                BodyExcel = BodyExcel + "<td> ShippingServiceLevel</td>";
                BodyExcel = BodyExcel + "<td> ShippedByUser</td>";
                BodyExcel = BodyExcel + "<td> ShippingWeight</td>";
                BodyExcel = BodyExcel + "<td> Weight</td>";
                BodyExcel = BodyExcel + "<td> Width</td>";
                BodyExcel = BodyExcel + "<td> Length</td>";
                BodyExcel = BodyExcel + "<td> Height</td>";
                BodyExcel = BodyExcel + "<td> StateRegion</td>";
                BodyExcel = BodyExcel + "<td> TrackingNum</td>";
                BodyExcel = BodyExcel + "<td> MfrName</td>";
                BodyExcel = BodyExcel + "<td> PricingRule</td>";

                BodyExcel = BodyExcel + "</tr>";

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";
            }

        }

        // realiza la impresion del cargo enviado si la tuviera el reporte de fedex
        // -------------------------------------------------------------------------------
        private void ColumnaCargo(string NombreCargo, string TrackingIDChargeDescription, string TrackingIDChargeAmount, string TrackingIDChargeDescription1, string TrackingIDChargeAmount1, string TrackingIDChargeDescription2, string TrackingIDChargeAmount2, string TrackingIDChargeDescription3, string TrackingIDChargeAmount3, string TrackingIDChargeDescription4, string TrackingIDChargeAmount4, string TrackingIDChargeDescription5, string TrackingIDChargeAmount5, string TrackingIDChargeDescription6, string TrackingIDChargeAmount6, string TrackingIDChargeDescription7, string TrackingIDChargeAmount7, string TrackingIDChargeDescription8, string TrackingIDChargeAmount8, string TrackingIDChargeDescription9, string TrackingIDChargeAmount9, string TrackingIDChargeDescription10, string TrackingIDChargeAmount10, string TrackingIDChargeDescription11, string TrackingIDChargeAmount11, string TrackingIDChargeDescription12, string TrackingIDChargeAmount12, string TrackingIDChargeDescription13, string TrackingIDChargeAmount13, string TrackingIDChargeDescription14, string TrackingIDChargeAmount14, string TrackingIDChargeDescription15, string TrackingIDChargeAmount15, string TrackingIDChargeDescription16, string TrackingIDChargeAmount16, string TrackingIDChargeDescription17, string TrackingIDChargeAmount17, string TrackingIDChargeDescription18, string TrackingIDChargeAmount18, string TrackingIDChargeDescription19, string TrackingIDChargeAmount19, string TrackingIDChargeDescription20, string TrackingIDChargeAmount20, string TrackingIDChargeDescription21, string TrackingIDChargeAmount21, string TrackingIDChargeDescription22, string TrackingIDChargeAmount22, string TrackingIDChargeDescription23, string TrackingIDChargeAmount23, string TrackingIDChargeDescription24, string TrackingIDChargeAmount24)
        {
            if (TrackingIDChargeDescription == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount + "</td>";
            else if (TrackingIDChargeDescription1 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount1 + "</td>";
            else if (TrackingIDChargeDescription2 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount2 + "</td>";
            else if (TrackingIDChargeDescription3 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount3 + "</td>";
            else if (TrackingIDChargeDescription4 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount4 + "</td>";
            else if (TrackingIDChargeDescription5 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount5 + "</td>";
            else if (TrackingIDChargeDescription6 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount6 + "</td>";
            else if (TrackingIDChargeDescription7 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount7 + "</td>";
            else if (TrackingIDChargeDescription8 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount8 + "</td>";
            else if (TrackingIDChargeDescription9 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount9 + "</td>";
            else if (TrackingIDChargeDescription10 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount10 + "</td>";
            else if (TrackingIDChargeDescription11 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount11 + "</td>";
            else if (TrackingIDChargeDescription12 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount12 + "</td>";
            else if (TrackingIDChargeDescription13 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount13 + "</td>";
            else if (TrackingIDChargeDescription14 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount14 + "</td>";
            else if (TrackingIDChargeDescription15 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount15 + "</td>";
            else if (TrackingIDChargeDescription16 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount16 + "</td>";
            else if (TrackingIDChargeDescription17 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount17 + "</td>";
            else if (TrackingIDChargeDescription18 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount18 + "</td>";
            else if (TrackingIDChargeDescription19 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount19 + "</td>";
            else if (TrackingIDChargeDescription20 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount20 + "</td>";
            else if (TrackingIDChargeDescription21 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount21 + "</td>";
            else if (TrackingIDChargeDescription22 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount22 + "</td>";
            else if (TrackingIDChargeDescription23 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount23 + "</td>";
            else if (TrackingIDChargeDescription24 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount24 + "</td>";
            else
                BodyExcel = BodyExcel + "<td>" + " " + "</td>";
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporte()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                var PedidoCollection = from s in listaPedido
                                       where s.FullTrakingId == TrackingNum
                                       select new
                                       {
                                           s.GroundTrackingIDPrefix,
                                           s.ExpressorGroundTrackingID,
                                           s.NetChargeAmount,
                                           s.ServiceType,
                                           s.GroundService,
                                           s.ShipmentDate,
                                           s.PODDeliveryDate,
                                           s.ActualWeightAmount,
                                           s.RatedWeightAmount,
                                           s.DimLength,
                                           s.DimWidth,
                                           s.DimHeight,
                                           s.DimDivisor,
                                           s.ShipperState,
                                           s.ZoneCode,
                                           s.TenderedDate,
                                           s.TrackingIDChargeDescription,
                                           s.TrackingIDChargeAmount,
                                           s.TrackingIDChargeDescription1,
                                           s.TrackingIDChargeAmount1,
                                           s.TrackingIDChargeDescription2,
                                           s.TrackingIDChargeAmount2,
                                           s.TrackingIDChargeDescription3,
                                           s.TrackingIDChargeAmount3,
                                           s.TrackingIDChargeDescription4,
                                           s.TrackingIDChargeAmount4,
                                           s.TrackingIDChargeDescription5,
                                           s.TrackingIDChargeAmount5,
                                           s.TrackingIDChargeDescription6,
                                           s.TrackingIDChargeAmount6,
                                           s.TrackingIDChargeDescription7,
                                           s.TrackingIDChargeAmount7,
                                           s.TrackingIDChargeDescription8,
                                           s.TrackingIDChargeAmount8,
                                           s.TrackingIDChargeDescription9,
                                           s.TrackingIDChargeAmount9,
                                           s.TrackingIDChargeDescription10,
                                           s.TrackingIDChargeAmount10,
                                           s.TrackingIDChargeDescription11,
                                           s.TrackingIDChargeAmount11,
                                           s.TrackingIDChargeDescription12,
                                           s.TrackingIDChargeAmount12,
                                           s.TrackingIDChargeDescription13,
                                           s.TrackingIDChargeAmount13,
                                           s.TrackingIDChargeDescription14,
                                           s.TrackingIDChargeAmount14,
                                           s.TrackingIDChargeDescription15,
                                           s.TrackingIDChargeAmount15,
                                           s.TrackingIDChargeDescription16,
                                           s.TrackingIDChargeAmount16,
                                           s.TrackingIDChargeDescription17,
                                           s.TrackingIDChargeAmount17,
                                           s.TrackingIDChargeDescription18,
                                           s.TrackingIDChargeAmount18,
                                           s.TrackingIDChargeDescription19,
                                           s.TrackingIDChargeAmount19,
                                           s.TrackingIDChargeDescription20,
                                           s.TrackingIDChargeAmount20,
                                           s.TrackingIDChargeDescription21,
                                           s.TrackingIDChargeAmount21,
                                           s.TrackingIDChargeDescription22,
                                           s.TrackingIDChargeAmount22,
                                           s.TrackingIDChargeDescription23,
                                           s.TrackingIDChargeAmount23,
                                           s.TrackingIDChargeDescription24,
                                           s.TrackingIDChargeAmount24
                                       };

                foreach (var Pedido in PedidoCollection)
                {

                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                    else
                        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                    // archivo base
                    // ------------
                    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                    // archivo fedex
                    // -------------
                    BodyExcel = BodyExcel + "<td>'" + Pedido.GroundTrackingIDPrefix + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + Pedido.ExpressorGroundTrackingID+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.NetChargeAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ServiceType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.GroundService + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ShipmentDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.PODDeliveryDate+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ActualWeightAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.RatedWeightAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimLength + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimDivisor + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ShipperState + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ZoneCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.TenderedDate + "</td>";

                    string NombreCargo = "Earned Discount";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Fuel Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Performance Pricing";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Delivery Area Surcharge Extended";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Delivery Area Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "USPS Non-Mach Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Residential";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Grace Discount";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Declared Value";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Extended Resi";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Additional Handling";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Parcel Re-Label Charge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Indirect Signature";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    //NombreCargo = "DAS Resi";
                    //ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Address Correction";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Extended Comm";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Oversize Charge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "AHS - Dimensions";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    BodyExcel = BodyExcel + "</tr>";
                    break;
                }

                FileExcel.WriteLine(BodyExcel);
                BodyExcel= "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporteUSPS()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

               //var PedidoCollection = from s in listaPedidoUSPS
               //                       where s.TrackingNumber == TrackingNum
               //                       select new
               //                       {
               //                            s.GroundService,
               //                            s.TrackingNumber,
               //                            s.NetChargeAmount,
               //                            s.PODDeliveryDate,
               //                            s.RatedWeightAmount,
               //                            s.ZoneCode
               //                       };

                //foreach (var Pedido in PedidoCollection)
                //{
                //    // arma la fila con el color de fondo que corresponde
                //    // --------------------------------------------------
                //    if (contador == 1)
                //        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                //    else
                //        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";
                //
                //    // archivo base
                //    // ------------
                //    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                //    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";
                //
                //    // archivo fedex
                //    // -------------
                //    string vacio = ""; 
                //    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //
                //    string NombreCargo = "Earned Discount";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Fuel Surcharge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Performance Pricing";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Delivery Area Surcharge Extended";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Delivery Area Surcharge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "USPS Non-Mach Surcharge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Residential";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Grace Discount";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Declared Value";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Extended Resi";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Additional Handling";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Parcel Re-Label Charge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Indirect Signature";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Resi";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Resi";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Address Correction";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Extended Comm";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Oversize Charge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "AHS - Dimensions";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    // dato USPS
                //    BodyExcel = BodyExcel + "<td>" + Pedido.GroundService + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.TrackingNumber + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.NetChargeAmount + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.PODDeliveryDate + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.RatedWeightAmount + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.ZoneCode + "</td>";
                //
                //    BodyExcel = BodyExcel + "</tr>";
                //    break;
                //}

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporteUPS()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                var PedidoCollection = from s in listaPedidoUPS
                                       where s.Campo30 == TrackingNum
                                       select new
                                       {
                                           s.Campo12,
                                           s.Campo30,
                                           s.Campo39
                                       };

                foreach (var Pedido in PedidoCollection)
                {
                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                    else
                        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                    // archivo base
                    // ------------
                    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                    // archivo fedex
                    // -------------
                    string vacio = "";
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";

                    string NombreCargo = "Earned Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Fuel Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Performance Pricing";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge Extended";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "USPS Non-Mach Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Residential";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Grace Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Declared Value";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Additional Handling";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Parcel Re-Label Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Indirect Signature";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Address Correction";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Comm";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Oversize Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "AHS - Dimensions";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    // dato USPS
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";

                    // dato UPS
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo12 + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo30 + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo39 + "</td>";

                    BodyExcel = BodyExcel + "</tr>";
                    break;
                }

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }


        //  inserta fila a reporte
        // -----------------------
        private void VentasNoFacturado()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoVentasNoFacturado, true))
            {
                // arma la fila con el color de fondo que corresponde
                // --------------------------------------------------
                if (contador == 1)
                    BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                else
                    BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                // archivo base
                // ------------
                BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                BodyExcel = BodyExcel + "</tr>";

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }


        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistro(string[] valor/*DataRow row*/)
        {
            SalesOrderNumber             = valor[0];//Convert.ToString(row["SalesOrderNumber"]);
            HoldCode                     = valor[1];//Convert.ToString(row["HoldCode"]);
            TotalSales                   = valor[2];//Convert.ToString(row["TotalSales"]);
            SalesSku                     = valor[3];//Convert.ToString(row["SalesSku"]);
            SalesCategoryAtTimeOfSale    = valor[4];//Convert.ToString(row["SalesCategoryAtTimeOfSale"]);
            UomCode                      = valor[5];//Convert.ToString(row["UomCode"]);
            UomQuantity                  = valor[6];//Convert.ToString(row["UomQuantity"]);
            SalesStatus                  = valor[7];//Convert.ToString(row["SalesStatus"]);
            SalesOrderDate               = valor[8];//Convert.ToString(row["SalesOrderDate"]);
            SalesChannelName             = valor[9];//Convert.ToString(row["SalesChannelName"]);
            CustomerName                 = valor[10];//Convert.ToString(row["CustomerName"]);
            FulfillmentSku               = valor[11];//Convert.ToString(row["FulfillmentSku"]);
            FulfillmentChannelName       = valor[12];//Convert.ToString(row["FulfillmentChannelName"]);
            FulfillmentChannelType       = valor[13];//Convert.ToString(row["FulfillmentChannelType"]);
            LinkedFulfillmentChannelName = valor[14];//Convert.ToString(row["LinkedFulfillmentChannelName"]);
            FulfillmentLocationName      = valor[15];//Convert.ToString(row["FulfillmentLocationName"]);
            FulfillmentOrderNumber       = valor[16];//Convert.ToString(row["FulfillmentOrderNumber"]);
            Quantity                     = valor[17];//Convert.ToString(row["Quantity"]);
            Sku                          = valor[18];//Convert.ToString(row["Sku"]);
            Title                        = valor[19];//Convert.ToString(row["Title"]);
            TotalCost                    = valor[20];//Convert.ToString(row["TotalCost"]);
            Commission                   = valor[21];//Convert.ToString(row["Commission"]);
            InventoryCost                = valor[22];//Convert.ToString(row["InventoryCost"]);
            UnitCost                     = valor[23];//Convert.ToString(row["UnitCost"]);
            ServiceCost                  = valor[24];//Convert.ToString(row["ServiceCost"]);
            EstimatedShippingCost        = valor[25];//Convert.ToString(row["EstimatedShippingCost"]);
            ShippingCost                 = valor[26];//Convert.ToString(row["ShippingCost"]);
            ShippingPrice                = valor[27];//Convert.ToString(row["ShippingPrice"]);
            OverheadCost                 = valor[28];//Convert.ToString(row["OverheadCost"]);
            PackageCost                  = valor[29];//Convert.ToString(row["PackageCost"]);
            ProfitLoss                   = valor[30];//Convert.ToString(row["ProfitLoss"]);
            Carrier                      = valor[31];//Convert.ToString(row["Carrier"]);
            ShippingServiceLevel         = valor[32];//Convert.ToString(row["ShippingServiceLevel"]);
            ShippedByUser                = valor[33];//Convert.ToString(row["ShippedByUser"]);
            ShippingWeight               = valor[34];//Convert.ToString(row["ShippingWeight"]);
            Length                       = valor[35];//Convert.ToString(row["Length"]);
            varWidth                     = valor[36];//Convert.ToString(row["Width"]);
            varHeight                    = valor[37];//Convert.ToString(row["Height"]);
            Weight                       = valor[38];//Convert.ToString(row["Weight"]);
            StateRegion                  = valor[39];//Convert.ToString(row["StateRegion"]);
            TrackingNum                  = valor[40];//Convert.ToString(row["TrackingNum"]);
            MfrName                      = valor[41];//Convert.ToString(row["MfrName"]);
            PricingRule                  = valor[42];//Convert.ToString(row["PricingRule"]);
            ActualShippingCost           = valor[43];
            ActualShipping               = valor[44];
            ShippingCostDifference       = valor[45];
        }
        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalle(DataRow row, ref PedidoFedex clsPedido)
        {
            clsPedido.GroundTrackingIDPrefix = Convert.ToString(row["Ground Tracking ID Prefix"]);
            clsPedido.ExpressorGroundTrackingID = Convert.ToString(row["Express or Ground Tracking ID"]);
            clsPedido.FullTrakingId = clsPedido.GroundTrackingIDPrefix + clsPedido.ExpressorGroundTrackingID;
            clsPedido.BilltoAccountNumber = Convert.ToString(row["Bill to Account Number"]);
            clsPedido.InvoiceDate = Convert.ToString(row["Invoice Date"]);
            clsPedido.InvoiceNumber = Convert.ToString(row["Invoice Number"]);
            clsPedido.StoreID = Convert.ToString(row["Store ID"]);
            clsPedido.OriginalAmountDue = Convert.ToString(row["Original Amount Due"]);
            clsPedido.CurrentBalance = Convert.ToString(row["Current Balance"]);
            clsPedido.Payor = Convert.ToString(row["Payor"]);
            clsPedido.TransportationChargeAmount = Convert.ToString(row["Transportation Charge Amount"]);
            clsPedido.NetChargeAmount = Convert.ToString(row["Net Charge Amount"]);
            clsPedido.ServiceType = Convert.ToString(row["Service Type"]);
            clsPedido.GroundService = Convert.ToString(row["Ground Service"]);
            clsPedido.ShipmentDate = Convert.ToString(row["Shipment Date"]);
            clsPedido.PODDeliveryDate = Convert.ToString(row["POD Delivery Date"]);
            clsPedido.PODDeliveryTime = Convert.ToString(row["POD Delivery Time"]);
            clsPedido.PODServiceAreaCode = Convert.ToString(row["POD Service Area Code"]);
            clsPedido.PODSignatureDescription = Convert.ToString(row["POD Signature Description"]);
            clsPedido.ActualWeightAmount = Convert.ToString(row["Actual Weight Amount"]);
            clsPedido.ActualWeightUnits = Convert.ToString(row["Actual Weight Units"]);
            clsPedido.RatedWeightAmount = Convert.ToString(row["Rated Weight Amount"]);
            clsPedido.RatedWeightUnits = Convert.ToString(row["Rated Weight Units"]);
            clsPedido.NumberofPieces = Convert.ToString(row["Number of Pieces"]);
            clsPedido.BundleNumber = Convert.ToString(row["Bundle Number"]);
            clsPedido.MeterNumber = Convert.ToString(row["Meter Number"]);
            clsPedido.TDMasterTrackingID = Convert.ToString(row["TDMasterTrackingID"]);
            clsPedido.ServicePackaging = Convert.ToString(row["Service Packaging"]);
            clsPedido.DimLength = Convert.ToString(row["Dim Length"]);
            clsPedido.DimWidth = Convert.ToString(row["Dim Width"]);
            clsPedido.DimHeight = Convert.ToString(row["Dim Height"]);
            clsPedido.DimDivisor = Convert.ToString(row["Dim Divisor"]);
            clsPedido.DimUnit = Convert.ToString(row["Dim Unit"]);
            clsPedido.RecipientName = Convert.ToString(row["Recipient Name"]);
            clsPedido.RecipientCompany = Convert.ToString(row["Recipient Company"]);
            clsPedido.RecipientAddressLine1 = Convert.ToString(row["Recipient Address Line 1"]);
            clsPedido.RecipientAddressLine2 = Convert.ToString(row["Recipient Address Line 2"]);
            clsPedido.RecipientCity = Convert.ToString(row["Recipient City"]);
            clsPedido.RecipientState = Convert.ToString(row["Recipient State"]);
            clsPedido.RecipientZipCode = Convert.ToString(row["Recipient Zip Code"]);
            clsPedido.ShipperCompany = Convert.ToString(row["Shipper Company"]);
            clsPedido.ShipperName = Convert.ToString(row["Shipper Name"]);
            clsPedido.ShipperAddressLine1 = Convert.ToString(row["Shipper Address Line 1"]);
            clsPedido.ShipperAddressLine2 = Convert.ToString(row["Shipper Address Line 2"]);
            clsPedido.ShipperCity = Convert.ToString(row["Shipper City"]);
            clsPedido.ShipperState = Convert.ToString(row["Shipper State"]);
            clsPedido.ShipperZipCode = Convert.ToString(row["Shipper Zip Code"]);
            clsPedido.OriginalCustomerReference = Convert.ToString(row["Original Customer Reference"]);
            clsPedido.OriginalDepartmentReferenceDescription = Convert.ToString(row["Original Department Reference Description"]);
            clsPedido.UpdatedCustomerReference = Convert.ToString(row["Updated Customer Reference"]);
            clsPedido.UpdatedDepartmentReferenceDescription = Convert.ToString(row["Updated Department Reference Description"]);
            clsPedido.OriginalRecipientAddressLine1 = Convert.ToString(row["Original Recipient Address Line 1"]);
            clsPedido.OriginalRecipientAddressLine2 = Convert.ToString(row["Original Recipient Address Line 2"]);
            clsPedido.OriginalRecipientCity = Convert.ToString(row["Original Recipient City"]);
            clsPedido.OriginalRecipientState = Convert.ToString(row["Original Recipient State"]);
            clsPedido.OriginalRecipientZipCode = Convert.ToString(row["Original Recipient Zip Code"]);
            clsPedido.ZoneCode = Convert.ToString(row["Zone Code"]);
            clsPedido.CostAllocation = Convert.ToString(row["Cost Allocation"]);
            clsPedido.AlternateAddressLine1 = Convert.ToString(row["Alternate Address Line 1"]);
            clsPedido.AlternateAddressLine2 = Convert.ToString(row["Alternate Address Line 2"]);
            clsPedido.AlternateCity = Convert.ToString(row["Alternate City"]);
            clsPedido.AlternateStateProvince = Convert.ToString(row["Alternate State Province"]);
            clsPedido.AlternateZipCode = Convert.ToString(row["Alternate Zip Code"]);
            clsPedido.CrossRefTrackingIDPrefix = Convert.ToString(row["CrossRefTrackingID Prefix"]);
            clsPedido.CrossRefTrackingID = Convert.ToString(row["CrossRefTrackingID"]);
            clsPedido.EntryDate = Convert.ToString(row["Entry Date"]);
            clsPedido.EntryNumber = Convert.ToString(row["Entry Number"]);
            clsPedido.CustomsValue = Convert.ToString(row["Customs Value"]);
            clsPedido.CustomsValueCurrencyCode = Convert.ToString(row["Customs Value Currency Code"]);
            clsPedido.DeclaredValue = Convert.ToString(row["Declared Value"]);
            clsPedido.DeclaredValueCurrencyCode = Convert.ToString(row["Declared Value Currency Code"]);
            clsPedido.CurrencyConversionDate = Convert.ToString(row["Currency Conversion Date"]);
            clsPedido.CurrencyConversionRate = Convert.ToString(row["Currency Conversion Rate"]);
            clsPedido.MultiweightNumber = Convert.ToString(row["Multiweight Number"]);
            clsPedido.MultiweightTotalMultiweightUnits = Convert.ToString(row["Multiweight Total Multiweight Units"]);
            clsPedido.MultiweightTotalMultiweightWeight = Convert.ToString(row["Multiweight Total Multiweight Weight"]);
            clsPedido.MultiweightTotalShipmentChargeAmount = Convert.ToString(row["Multiweight Total Shipment Charge Amount"]);
            clsPedido.MultiweightTotalShipmentWeight = Convert.ToString(row["Multiweight Total Shipment Weight"]);
            clsPedido.GroundTrackingIDAddressCorrectionDiscountChargeAmount = Convert.ToString(row["Ground Tracking ID Address Correction Discount Charge Amount"]);
            clsPedido.GroundTrackingIDAddressCorrectionGrossChargeAmount = Convert.ToString(row["Ground Tracking ID Address Correction Gross Charge Amount"]);
            clsPedido.RatedMethod = Convert.ToString(row["Rated Method"]);
            clsPedido.SortHub = Convert.ToString(row["Sort Hub"]);
            clsPedido.EstimatedWeight = Convert.ToString(row["Estimated Weight"]);
            clsPedido.EstimatedWeightUnit = Convert.ToString(row["Estimated Weight Unit"]);
            clsPedido.PostalClass = Convert.ToString(row["Postal Class"]);
            clsPedido.ProcessCategory = Convert.ToString(row["Process Category"]);
            clsPedido.PackageSize = Convert.ToString(row["Package Size"]);
            clsPedido.DeliveryConfirmation = Convert.ToString(row["Delivery Confirmation"]);
            clsPedido.TenderedDate = Convert.ToString(row["Tendered Date"]);
            clsPedido.TrackingIDChargeDescription = Convert.ToString(row["Tracking ID Charge Description"]);
            clsPedido.TrackingIDChargeAmount = Convert.ToString(row["Tracking ID Charge Amount"]);
            clsPedido.TrackingIDChargeDescription1 = Convert.ToString(row["Tracking ID Charge Description1"]);
            clsPedido.TrackingIDChargeAmount1 = Convert.ToString(row["Tracking ID Charge Amount1"]);
            clsPedido.TrackingIDChargeDescription2 = Convert.ToString(row["Tracking ID Charge Description2"]);
            clsPedido.TrackingIDChargeAmount2 = Convert.ToString(row["Tracking ID Charge Amount2"]);
            clsPedido.TrackingIDChargeDescription3 = Convert.ToString(row["Tracking ID Charge Description3"]);
            clsPedido.TrackingIDChargeAmount3 = Convert.ToString(row["Tracking ID Charge Amount3"]);
            clsPedido.TrackingIDChargeDescription4 = Convert.ToString(row["Tracking ID Charge Description4"]);
            clsPedido.TrackingIDChargeAmount4 = Convert.ToString(row["Tracking ID Charge Amount4"]);
            clsPedido.TrackingIDChargeDescription5 = Convert.ToString(row["Tracking ID Charge Description5"]);
            clsPedido.TrackingIDChargeAmount5 = Convert.ToString(row["Tracking ID Charge Amount5"]);
            clsPedido.TrackingIDChargeDescription6 = Convert.ToString(row["Tracking ID Charge Description6"]);
            clsPedido.TrackingIDChargeAmount6 = Convert.ToString(row["Tracking ID Charge Amount6"]);
            clsPedido.TrackingIDChargeDescription7 = Convert.ToString(row["Tracking ID Charge Description7"]);
            clsPedido.TrackingIDChargeAmount7 = Convert.ToString(row["Tracking ID Charge Amount7"]);
            clsPedido.TrackingIDChargeDescription8 = Convert.ToString(row["Tracking ID Charge Description8"]);
            clsPedido.TrackingIDChargeAmount8 = Convert.ToString(row["Tracking ID Charge Amount8"]);
            clsPedido.TrackingIDChargeDescription9 = Convert.ToString(row["Tracking ID Charge Description9"]);
            clsPedido.TrackingIDChargeAmount9 = Convert.ToString(row["Tracking ID Charge Amount9"]);
            clsPedido.TrackingIDChargeDescription10 = Convert.ToString(row["Tracking ID Charge Description10"]);
            clsPedido.TrackingIDChargeAmount10 = Convert.ToString(row["Tracking ID Charge Amount10"]);
            clsPedido.TrackingIDChargeDescription11 = Convert.ToString(row["Tracking ID Charge Description11"]);
            clsPedido.TrackingIDChargeAmount11 = Convert.ToString(row["Tracking ID Charge Amount11"]);
            clsPedido.TrackingIDChargeDescription12 = Convert.ToString(row["Tracking ID Charge Description12"]);
            clsPedido.TrackingIDChargeAmount12 = Convert.ToString(row["Tracking ID Charge Amount12"]);
            clsPedido.TrackingIDChargeDescription13 = Convert.ToString(row["Tracking ID Charge Description13"]);
            clsPedido.TrackingIDChargeAmount13 = Convert.ToString(row["Tracking ID Charge Amount13"]);
            clsPedido.TrackingIDChargeDescription14 = Convert.ToString(row["Tracking ID Charge Description14"]);
            clsPedido.TrackingIDChargeAmount14 = Convert.ToString(row["Tracking ID Charge Amount14"]);
            clsPedido.TrackingIDChargeDescription15 = Convert.ToString(row["Tracking ID Charge Description15"]);
            clsPedido.TrackingIDChargeAmount15 = Convert.ToString(row["Tracking ID Charge Amount15"]);
            clsPedido.TrackingIDChargeDescription16 = Convert.ToString(row["Tracking ID Charge Description16"]);
            clsPedido.TrackingIDChargeAmount16 = Convert.ToString(row["Tracking ID Charge Amount16"]);
            clsPedido.TrackingIDChargeDescription17 = Convert.ToString(row["Tracking ID Charge Description17"]);
            clsPedido.TrackingIDChargeAmount17 = Convert.ToString(row["Tracking ID Charge Amount17"]);
            clsPedido.TrackingIDChargeDescription18 = Convert.ToString(row["Tracking ID Charge Description18"]);
            clsPedido.TrackingIDChargeAmount18 = Convert.ToString(row["Tracking ID Charge Amount18"]);
            clsPedido.TrackingIDChargeDescription19 = Convert.ToString(row["Tracking ID Charge Description19"]);
            clsPedido.TrackingIDChargeAmount19 = Convert.ToString(row["Tracking ID Charge Amount19"]);
            clsPedido.TrackingIDChargeDescription20 = Convert.ToString(row["Tracking ID Charge Description20"]);
            clsPedido.TrackingIDChargeAmount20 = Convert.ToString(row["Tracking ID Charge Amount20"]);
            clsPedido.TrackingIDChargeDescription21 = Convert.ToString(row["Tracking ID Charge Description21"]);
            clsPedido.TrackingIDChargeAmount21 = Convert.ToString(row["Tracking ID Charge Amount21"]);
            clsPedido.TrackingIDChargeDescription22 = Convert.ToString(row["Tracking ID Charge Description22"]);
            clsPedido.TrackingIDChargeAmount22 = Convert.ToString(row["Tracking ID Charge Amount22"]);
            clsPedido.TrackingIDChargeDescription23 = Convert.ToString(row["Tracking ID Charge Description23"]);
            clsPedido.TrackingIDChargeAmount23 = Convert.ToString(row["Tracking ID Charge Amount23"]);
            clsPedido.TrackingIDChargeDescription24 = Convert.ToString(row["Tracking ID Charge Description24"]);
            clsPedido.TrackingIDChargeAmount24 = Convert.ToString(row["Tracking ID Charge Amount24"]);
            clsPedido.ShipmentNotes = Convert.ToString(row["Shipment Notes"]);
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleUSPS(DataRow row, ref PedidoUSPS clsPedido )
        {
            clsPedido.AccountNumber = Convert.ToString(row["Account Number"]);
            clsPedido.ID = Convert.ToString(row["ID"]);
            //clsPedido.DateTime = Convert.ToString(row["Date/Time"]);
            clsPedido.Postmark = Convert.ToString(row["Postmark"]);
            clsPedido.Origin = Convert.ToString(row["Origin"]);
            clsPedido.Destination = Convert.ToString(row["Destination"]);
            clsPedido.Type = Convert.ToString(row["Type"]);
            clsPedido.MailClass = Convert.ToString(row["Mail Class"]);
            clsPedido.TrackingNumber = Convert.ToString(row["Tracking Number"]);
            clsPedido.DeclaredValue = Convert.ToString(row["Declared Value"]);
            clsPedido.TotalPostageAmt = Convert.ToString(row["Total Postage Amt"]);
            clsPedido.Balance = Convert.ToString(row["Balance"]);
            clsPedido.RefundStatus = Convert.ToString(row["Refund Status"]);
            clsPedido.GroupCode = Convert.ToString(row["Group Code"]);
            clsPedido.ReferenceID = Convert.ToString(row["Reference ID"]);
            clsPedido.DeliveryDate = Convert.ToString(row["Delivery Date"]);
            clsPedido.StatusCode = Convert.ToString(row["Status Code"]);
            clsPedido.StatusDescription = Convert.ToString(row["Status Description"]);
            clsPedido.Weight = Convert.ToString(row["Weight"]);
            clsPedido.OptionalServices = Convert.ToString(row["OptionalServices"]);
            clsPedido.DestinationName = Convert.ToString(row["Destination Name"]);
            clsPedido.DestinationCompanyName = Convert.ToString(row["Destination Company Name"]);
            clsPedido.DestinationAddress = Convert.ToString(row["Destination Address"]);
            clsPedido.DestinationCity = Convert.ToString(row["Destination City"]);
            clsPedido.DestinationState = Convert.ToString(row["Destination State"]);
            clsPedido.DestinationZip = Convert.ToString(row["Destination Zip"]);
            clsPedido.DestinationCountry = Convert.ToString(row["Destination Country"]);
            clsPedido.Phone = Convert.ToString(row["Phone"]);
            clsPedido.Email = Convert.ToString(row["Email"]);
            clsPedido.Reference2 = Convert.ToString(row["Reference2"]);
            clsPedido.Reference3 = Convert.ToString(row["Reference3"]);
            clsPedido.Reference4 = Convert.ToString(row["Reference4"]);
            clsPedido.PackageDescription = Convert.ToString(row["Package Description"]);
            clsPedido.Zone = Convert.ToString(row["Zone"]);
            clsPedido.IsCubic = Convert.ToString(row["IsCubic"]);
            clsPedido.CubicValue = Convert.ToString(row["Cubic Value"]);
            //clsPedido.AdjWeight = Convert.ToString(row["Adj. Weight"]);
            //clsPedido.AdjDimensions = Convert.ToString(row["Adj. Dimensions"]);
            //clsPedido.AdjFromZIP = Convert.ToString(row["Adj. From ZIP"]);
            //clsPedido.AdjToZIP = Convert.ToString(row["Adj. To ZIP"]);
            //clsPedido.AdjMailClass = Convert.ToString(row["Adj. Mail Class"]);

        }

    // obtiene el valor del registro actual
    // ------------------------------------
    private void ObtieneValorRegistroDetalleUPS(string[] valor,/*DataRow row, */ref PedidoUPS clsPedido)
        {
            clsPedido.Campo1 = valor[0];
            clsPedido.Campo2 = valor[1];
            clsPedido.Campo3 = valor[2];
            clsPedido.Campo4 = valor[3];
            clsPedido.Campo5 = valor[4];
            clsPedido.Campo6 = valor[5];
            clsPedido.Campo7 = valor[6];
            clsPedido.Campo8 = valor[7];
            clsPedido.Campo9 = valor[8];
            clsPedido.Campo10 = valor[9];
            clsPedido.Campo11 = valor[10];
            clsPedido.Campo12 = valor[11];
            clsPedido.Campo13 = valor[12];
            clsPedido.Campo14 = valor[13];
            clsPedido.Campo15 = valor[14];
            clsPedido.Campo16 = valor[15];
            clsPedido.Campo17 = valor[16];
            clsPedido.Campo18 = valor[17];
            clsPedido.Campo19 = valor[18];
            clsPedido.Campo20 = valor[19];
            clsPedido.Campo21 = valor[20];
            clsPedido.Campo22 = valor[21];
            clsPedido.Campo23 = valor[22];
            clsPedido.Campo24 = valor[23];
            clsPedido.Campo25 = valor[24];
            clsPedido.Campo26 = valor[25];
            clsPedido.Campo27 = valor[26];
            clsPedido.Campo28 = valor[27];
            clsPedido.Campo29 = valor[28];
            clsPedido.Campo30 = valor[29];
            clsPedido.Campo31 = valor[30];
            clsPedido.Campo32 = valor[31];
            clsPedido.Campo33 = valor[32];
            clsPedido.Campo34 = valor[33];
            clsPedido.Campo35 = valor[34];
            clsPedido.Campo36 = valor[35];
            clsPedido.Campo37 = valor[36];
            clsPedido.Campo38 = valor[37];
            clsPedido.Campo39 = valor[38];
            clsPedido.Campo40 = valor[39];
            clsPedido.Campo41 = valor[40];
            clsPedido.Campo42 = valor[41];
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleAmazon(string[] valor, ref PedidoAmazon clsPedido)
        {
            clsPedido.date = valor[0];
            clsPedido.datetime = valor[1];
            clsPedido.settlementid = valor[2];
            clsPedido.type = valor[3];
            clsPedido.orderid = valor[4];
            clsPedido.sku = valor[5];
            clsPedido.description = valor[6];
            clsPedido.quantity = valor[7];
            clsPedido.marketplace = valor[8];
            clsPedido.fulfillment = valor[9];
            clsPedido.ordercity = valor[10];
            clsPedido.orderstate = valor[11];
            clsPedido.orderpostal = valor[12];
            clsPedido.taxcollectionmodel = valor[13];
            clsPedido.productsales = valor[14];
            clsPedido.productsalestax = valor[15];
            clsPedido.shippingcredits = valor[16];
            clsPedido.shippingcreditstax = valor[17];
            clsPedido.giftwrapcredits = valor[18];
            clsPedido.giftwrapcreditstax = valor[19];
            clsPedido.promotionalrebates = valor[20];
            clsPedido.promotionalrebatestax = valor[21];
            clsPedido.marketplacewithheldtax = valor[22];
            clsPedido.sellingfees = valor[23];
            clsPedido.fbafees = valor[24];
            clsPedido.othertransactionfees = valor[25];
            clsPedido.other = valor[26];
            clsPedido.total = valor[27];

        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleMI15(string[] valor, ref ShippingByMarket.MI15 clsPedido)
        {
            clsPedido.SHIPPINGDATE = valor[0];
            clsPedido.MANIFESTDATE = valor[1];
            clsPedido.PACKAGEID = valor[2];
            clsPedido.USPSTRACKINGNUMBER = valor[3];
            clsPedido.SEQUENCE = valor[4];
            clsPedido.COSTCENTER1 = valor[5];
            clsPedido.COSTCENTER2 = valor[6];
            clsPedido.COSTCENTER3 = valor[7];
            clsPedido.BILLEDWEIGHT = valor[8];
            clsPedido.WEIGHTTYPE = valor[9];
            clsPedido.ZIP = valor[10];
            clsPedido.ZONE = valor[11];
            clsPedido.SERVICE = valor[12];
            clsPedido.UPSMI = valor[13];
            clsPedido.USPS = valor[14];
            clsPedido.SAVINGS = valor[15];
            clsPedido.OVERLABELEDUSPSTRACKING = valor[16];
            clsPedido.ERRORREASON = valor[17];

        }

        // obtiene el valor del registro de box
        // ------------------------------------
        private void ObtieneValorRegistroDetalleBOX(DataRow row, ref ShippingByMarket.Box clsPedido)
        {
            clsPedido.STATE = Convert.ToString(row["STATE"]);
            clsPedido.POSTALCODE = Convert.ToString(row["POSTALCODE"]);
            clsPedido.SHIPPER = Convert.ToString(row["SHIPPER"]);
            clsPedido.PROSHIP_SHIPDATE = Convert.ToString(row["PROSHIP_SHIPDATE"]);
            clsPedido.PACKAGING_PLAINTEXT = Convert.ToString(row["PACKAGING_PLAINTEXT"]);
            clsPedido.WEIGHT = Convert.ToString(row["WEIGHT"]);
            clsPedido.DIMENSIONS = Convert.ToString(row["DIMENSIONS"]);
            clsPedido.TRACKING_NUMBER = Convert.ToString(row["TRACKING_NUMBER"]);
            clsPedido.CCN_SAP_ORDER_NUMBER = Convert.ToString(row["CCN_SAP_ORDER_NUMBER"]);
            clsPedido.CCN_ORDER_NUMBER = Convert.ToString(row["CCN_ORDER_NUMBER"]);
            clsPedido.CCN_COMPANY_CODE = Convert.ToString(row["CCN_COMPANY_CODE"]);
            clsPedido.CCN_STR_NUM = Convert.ToString(row["CCN_STR_NUM"]);
            clsPedido.CCN_DELIVERY_NUMBER = Convert.ToString(row["CCN_DELIVERY_NUMBER"]);
            clsPedido.SHIPPER_SYMBOL = Convert.ToString(row["SHIPPER_SYMBOL"]);
            clsPedido.OrderDate = Convert.ToString(row["Order Date"]);
            clsPedido.PROSHIP_SERVICE_PLAINTEXT = Convert.ToString(row["PROSHIP_SERVICE_PLAINTEXT"]);
            clsPedido.CCN_SHIP_TEXT = Convert.ToString(row["CCN_SHIP_TEXT"]);
        }

        // obtiene el valor del registro de EJDDimensions
        // ----------------------------------------------
        private void ObtieneValorRegistroDetalleEJDDimensions(DataRow row, ref ShippingByMarket.EJDDimensions clsPedido)
        {
            clsPedido.EvpSku = Convert.ToString(row["Evp Sku"]);
            clsPedido.Title = Convert.ToString(row["Title"]);
            clsPedido.EJDSku = Convert.ToString(row["EJD Sku"]);
            clsPedido.EJDUomCode = Convert.ToString(row["EJD Uom Code"]);
            clsPedido.EJDUomQuantity = Convert.ToString(row["EJD Uom Quantity"]);
            clsPedido.Length = Convert.ToString(row["Length"]);
            clsPedido.Height = Convert.ToString(row["Height"]);
            clsPedido.Width = Convert.ToString(row["Width"]);
            clsPedido.Weight = Convert.ToString(row["Weight"]);
        }

        // obtiene el valor del registro de EJDDimensions
        // ----------------------------------------------
        private void ObtieneValorRegistroDetalleJensenDimensions(DataRow row, ref ShippingByMarket.JensenDimensions clsPedido)
        {
            clsPedido.EvpSku = Convert.ToString(row["Evp Sku"]);
            clsPedido.Title = Convert.ToString(row["Title"]);
            clsPedido.JensenSku = Convert.ToString(row["Jensen Sku"]);
            clsPedido.UomCode = Convert.ToString(row["UomCode"]);
            clsPedido.UomQuantity = Convert.ToString(row["UomQuantity"]);
            clsPedido.Length = Convert.ToString(row["Length"]);
            clsPedido.Height = Convert.ToString(row["Height"]);
            clsPedido.Width = Convert.ToString(row["Width"]);
            clsPedido.Weight = Convert.ToString(row["Weight"]);
        }

        // obtiene el valor del registro de EJDDimensions
        // ----------------------------------------------
        private void ObtieneValorRegistroCancelados(DataRow row, ref ShippingByMarket.Clases.Cancelados clsPedido)
        {
            clsPedido.OrderDate             = Convert.ToDateTime(row["Order Date"]);
            clsPedido.PONumber              = Convert.ToString(row["PO Number"]);
            clsPedido.Status                = Convert.ToString(row["Status"]);
            clsPedido.Notes                 = Convert.ToString(row["Notes"]);
            clsPedido.Supplier              = Convert.ToString(row["Supplier"]);
            clsPedido.SupplierNumber        = Convert.ToString(row["Supplier Number"]);
            clsPedido.SupplierStatus         = Convert.ToString(row["Supplier Status"]);
            //clsPedido.Vacia                 = Convert.ToString(row["Supplier Status"]);
            clsPedido.ShipmentCount         = Convert.ToString(row["Shipment Count"]);
            clsPedido.Type                  = Convert.ToString(row["Type"]);
            clsPedido.PurchaseLocations     = Convert.ToString(row["Purchase Locations"]);
            clsPedido.ReceiveLocations      = Convert.ToString(row["Receive Locations"]);
            clsPedido.ItemSummary           = Convert.ToString(row["ItemSummary"]);
            clsPedido.ShippingServiceLevel  = Convert.ToString(row["ShippingServiceLevel"]);
            clsPedido.ShipTo                = Convert.ToString(row["Ship To"]);
            clsPedido.City                  = Convert.ToString(row["City"]);
            clsPedido.State                 = Convert.ToString(row["State"]);
            clsPedido.Country               = Convert.ToString(row["Country"]);
            clsPedido.PostalCode            = Convert.ToString(row["Postal Code"]);
            clsPedido.TotalWeight           = (float)Convert.ToDouble(row["TotalWeight"]);
            clsPedido.CreatedDate           = Convert.ToDateTime(row["Created Date"]);
            clsPedido.ExpectedDate          = Convert.ToDateTime(row["Expected Date"]);
            clsPedido.Total                 = (float)Convert.ToDouble(row["Total"]);
            clsPedido.FechaInsercion        = DateTime.Today; 
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosFedex(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroFEDEX(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString+@"\"+ ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios Fedex "+ DateTime.Now.ToString("yyyyMMddTHHmmss")+"\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }

 
            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: "+ Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }
                FlgSihayFedex = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name+" "+ DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PedidoFedex clsPedido = new PedidoFedex();
                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalle(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDFEDEX(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: "+ Archivos.Name+" " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString +@"\"+ Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos Fedex";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosUSPS(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroUSPS(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios USPS " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }



            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayUSPS = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PedidoUSPS clsPedido = new PedidoUSPS();
                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleUSPS(row, ref clsPedido);

                    // Valida no insertar registros de footer del archivo
                    // --------------------------------------------------
                    if (clsPedido.AccountNumber == "" || clsPedido.AccountNumber == "Transaction Type" || clsPedido.AccountNumber == "Adjustment" || clsPedido.AccountNumber == "Postage Print" || clsPedido.AccountNumber == "Postage Purchase" || clsPedido.AccountNumber == "Postage Refund")
                        continue;

                    clsInsertaRegistro.InsertaBDUSPS( clsPedido, Conexion);
                    //listaPedidoUSPS.Add(clsPedido);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos USPS";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosUPS(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroUPS(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios UPS " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienMI15"];
                //if (filemaster != "")
                //{
                //    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                //        continue;
                //}

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUPS"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayUPS = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader upsFile = new System.IO.StreamReader(file);
                string[] PalabraUps = new string[42];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = upsFile.ReadLine()) != null)
                {
                    string[] sa = line.Split(',');

                    if (sa.Length > 42)
                    {
                        //cantidad = sa.Length - 42;
                        continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        PalabraUps[contPala] = s;
                        contPala = contPala + 1;
                    }
                    
                    PedidoUPS clsPedido = new PedidoUPS();
                    ObtieneValorRegistroDetalleUPS(PalabraUps, ref clsPedido);
                    //listaPedidoUPS.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDUPS(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                upsFile.Close();

                // obtiene record set
                // ------------------
                //string connectionString = GetConnectionString(file, "DETALLE");
                //var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                //int conteoregistros = 0;
                //
                //// recorre registros obtenidos por la lectura del excel
                //// ----------------------------------------------------
                //foreach (DataRow row in dataSet.Tables[0].Rows)
                //{
                //    PedidoUPS clsPedido = new PedidoUPS();
                //    // obtiene el valor del registro leido
                //    // -----------------------------------
                //    ObtieneValorRegistroDetalleUPS(row, ref clsPedido);
                //
                //    listaPedidoUPS.Add(clsPedido);
                //    conteoregistros += 1;
                //}

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos UPS";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleEndicia(string[] valor, ref PedidoEndicia clsPedido)
        {
            clsPedido.PrintDate = Convert.ToDateTime(valor[0]);
            clsPedido.AmountPaid = System.Convert.ToDecimal(valor[1]);
            clsPedido.AdjAmount = valor[2];
            clsPedido.QuotedAmount = System.Convert.ToDecimal(valor[3]);
            clsPedido.Recipient = valor[4];
            clsPedido.Status = valor[5];
            clsPedido.TrackingNumber = valor[6];

            clsPedido.TrackingNumber = clsPedido.TrackingNumber.Replace("=", "");

            if (valor[7].Contains("/"))
                clsPedido.DateDelivered = Convert.ToDateTime(valor[7]);
            else
                clsPedido.DateDelivered = new DateTime(1000, 1, 1);

            clsPedido.Carrier = valor[8];
            clsPedido.ClassService = valor[9];
            clsPedido.InsuredValue = System.Convert.ToDecimal(valor[10]);
            clsPedido.InsuranceID = valor[11];
            clsPedido.CostCode = valor[12];
            clsPedido.Weight = valor[13];

            if (valor[14].Contains("/"))
                clsPedido.ShipDate = Convert.ToDateTime(valor[14]);
            else
                clsPedido.ShipDate = new DateTime(1000, 1, 1);

            clsPedido.RefundType = valor[15];
            clsPedido.PrintedMessage = valor[16];
            clsPedido.User = valor[17];

            if (valor[18].Contains("/"))
                clsPedido.RefundRequestDate = Convert.ToDateTime(valor[18]);
            else
                clsPedido.RefundRequestDate = new DateTime(1000, 1, 1);

            clsPedido.RefundStatus = valor[19];
            clsPedido.RefundRequested = valor[20];
            clsPedido.Reference1 = valor[21];
            clsPedido.Reference2 = valor[22];
            clsPedido.Reference3 = valor[23];
            clsPedido.Reference4 = valor[24];

        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosAmazon(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroAmazon(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios amazon " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            int ContadorArchivos = 0;
            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();


                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienAmazonOriginal"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayAmazon = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }


                // elimina las primeras 7 lineas ya que son encabezado de amazon
                int Contadorlineas = 0;
                string RutaAmazon = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
                RutaAmazon = RutaAmazon + "Amazon_" + ContadorArchivos.ToString() + ".csv";

                using (StreamWriter fileWrite = new StreamWriter(RutaAmazon))
                {
                    using (StreamReader fielRead = new StreamReader(Archivos.FullName))
                    {
                        String line;

                        while ((line = fielRead.ReadLine()) != null)
                        {
                            if (Contadorlineas <= 6)
                            {
                                Contadorlineas = Contadorlineas + 1;
                                continue;
                            }
                            else
                                fileWrite.WriteLine(line);
                        }

                        fielRead.Close();
                    }

                    fileWrite.Close();
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader amazonFile = new System.IO.StreamReader(RutaAmazon);

                string[] PalabraAmazon = new string[28];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = amazonFile.ReadLine()) != null)
                {
                    string LineaSinCaracteres = line.Replace("\"", "");
                    LineaSinCaracteres = LineaSinCaracteres.Replace("/", ",");
                    string[] sa = LineaSinCaracteres.Split(',');

                    if (sa.Length > 28)
                    {
                        //cantidad = sa.Length - 42;
                        continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        PalabraAmazon[contPala] = s;
                        contPala = contPala + 1;
                    }

                    if (PalabraAmazon[3] != "Shipping Services")
                        continue;

                    if (PalabraAmazon[6] != "Shipping Label Purchased through Amazon")
                    {
                        if(PalabraAmazon[6] == "Shipping Label Refunded through Amazon")
                        {
                            PedidoAmazon clsPedidorefund = new PedidoAmazon();
                            ObtieneValorRegistroDetalleAmazon(PalabraAmazon, ref clsPedidorefund);
                            //listaPedidoAmazon.Add(clsPedido);
                            clsInsertaRegistro.InsertaBDAMAZON(clsPedidorefund, Conexion,true);
                            conteoregistros += 1;
                        }

                        continue;
                    }

                    PedidoAmazon clsPedido = new PedidoAmazon();
                    ObtieneValorRegistroDetalleAmazon(PalabraAmazon, ref clsPedido);
                    //listaPedidoAmazon.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDAMAZON(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                amazonFile.Close();

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                string RutaArchivoMoverAmazon = pathString + @"\" + "Amazon_" + ContadorArchivos.ToString() + ".csv"; ;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);
                System.IO.File.Move(RutaAmazon, RutaArchivoMoverAmazon);
                ContadorArchivos += 1;

            }

            //textBox1.Text = "Fin Carga Archivos Amazon";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosEndicia(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroEndicia(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios Endicia " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }



            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienEndicia"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayEndicia = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader EndiciaFile = new System.IO.StreamReader(file);
                string[] PalabraEndicia = new string[26];
                int contPala = 0;
                int conteoregistros = 0;
                int Inicio = 0;
                string PalabraSinComa = "";
                string PalabraFinal = "";

                while ((line = EndiciaFile.ReadLine()) != null)
                {
                    PalabraFinal = "";
                    string NuevaLinea = line.Replace(@"""""", "");

                    if (NuevaLinea.Contains(@""""))
                    {
                        int Posicion = NuevaLinea.IndexOf(@"""", Inicio);

                        int PosicionFinal = NuevaLinea.IndexOf(@"""", Posicion + 1);

                        for (int ii = 0; ii < NuevaLinea.Length; ii++)
                        {
                            if (ii >= Posicion && ii <= PosicionFinal)
                            {
                                PalabraSinComa = NuevaLinea.Substring(ii, 1);

                                if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                {
                                    continue;
                                }
                                else
                                {

                                    PalabraFinal = PalabraFinal + NuevaLinea.Substring(ii, 1);
                                }
                            }
                            else
                            {
                                PalabraFinal = PalabraFinal + NuevaLinea.Substring(ii, 1);
                            }
                        }
                    }
                    else
                    {
                        PalabraFinal = line;
                    }

                    // busca siguiente coincidencia de ,
                    // ---------------------------------
                    string PalabraFinal2 = "";
                    if (PalabraFinal.Contains(@""""))
                    {
                        int Posicion = PalabraFinal.IndexOf(@"""", Inicio);

                        int PosicionFinal = PalabraFinal.IndexOf(@"""", Posicion + 1);

                        for (int ii = 0; ii < PalabraFinal.Length; ii++)
                        {
                            if (ii >= Posicion && ii <= PosicionFinal)
                            {
                                PalabraSinComa = PalabraFinal.Substring(ii, 1);

                                if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                {
                                    continue;
                                }
                                else
                                {

                                    PalabraFinal2 = PalabraFinal2 + PalabraFinal.Substring(ii, 1);
                                }
                            }
                            else
                            {
                                PalabraFinal2 = PalabraFinal2 + PalabraFinal.Substring(ii, 1);
                            }
                        }
                    }
                    else
                    {
                        PalabraFinal2 = PalabraFinal;
                    }

                    PalabraFinal2 = PalabraFinal2.Replace("$", "");
                    PalabraFinal2 = PalabraFinal2.Replace("\"", "");
                    string[] sa = PalabraFinal2.Split(',');

                    if (sa.Length > 26)
                    {
                        continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        PalabraEndicia[contPala] = s;
                        contPala = contPala + 1;
                    }

                    if (line.Contains("Print Date"))
                        continue;

                    PedidoEndicia clsPedido = new PedidoEndicia();
                    ObtieneValorRegistroDetalleEndicia(PalabraEndicia, ref clsPedido);
                    clsInsertaRegistro.InsertaBDEndicia(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                EndiciaFile.Close();

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos Endicia";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void CargaDatosBOX(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            //clsInsertaRegistro.EliminaRegistroBOX(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos BOX " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["ArchivoContieneBOX"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    ShippingByMarket.Box clsPedido = new ShippingByMarket.Box();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleBOX(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaBOX(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos Fedex";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void CargaDatosEJDDimensions(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroEJDDimensions(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------

            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos EJDDimensions " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["ArchivoContieneEJDDimensions"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    ShippingByMarket.EJDDimensions clsPedido = new ShippingByMarket.EJDDimensions();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleEJDDimensions(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaEJDDimensions(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos EJD";
            this.Refresh();
            this.Invalidate();
        }


        // Cargadatos excel Jensen
        // -----------------------
        private void CargaDatosJensenDimensions(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroJensenDimensions(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------

            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos JensenDimensions " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["ArchivoContieneJensenDimensions"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    ShippingByMarket.JensenDimensions clsPedido = new ShippingByMarket.JensenDimensions();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleJensenDimensions(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaJensenDimensions(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos JensenDimensions";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosMI15(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            //clsInsertaRegistro.EliminaRegistroM15(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios amazon " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            int ContadorArchivos = 0;
            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();


                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienAmazonOriginal"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienMI15"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ".csv";
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayMI15 = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }


                // elimina las primeras 7 lineas ya que son encabezado de amazon
                int Contadorlineas = 0;
                string RutaAmazon = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
                RutaAmazon = RutaAmazon + "MI15Final_" + ContadorArchivos.ToString() + ".csv";

                using (StreamWriter fileWrite = new StreamWriter(RutaAmazon))
                {
                    using (StreamReader fielRead = new StreamReader(Archivos.FullName))
                    {
                        String line;

                        while ((line = fielRead.ReadLine()) != null)
                        {
                            if (Contadorlineas <= 4)
                            {
                                Contadorlineas = Contadorlineas + 1;
                                continue;
                            }
                            else
                                fileWrite.WriteLine(line);
                        }

                        fielRead.Close();
                    }

                    fileWrite.Close();
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader MI15File = new System.IO.StreamReader(RutaAmazon);

                string[] PalabraMI15 = new string[18];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = MI15File.ReadLine()) != null)
                {
                    string LineaSinCaracteres = line.Replace("\"", "");
                    //LineaSinCaracteres = LineaSinCaracteres.Replace("/", ",");
                    string[] sa = LineaSinCaracteres.Split(',');

                    if (sa.Length > 18)
                    {
                        //cantidad = sa.Length - 42;
                        continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        PalabraMI15[contPala] = s;
                        contPala = contPala + 1;
                    }


                    ShippingByMarket.MI15 clsPedido = new ShippingByMarket.MI15();
                    ObtieneValorRegistroDetalleMI15(PalabraMI15, ref clsPedido);
                    clsInsertaRegistro.InsertaMI15(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                MI15File.Close();

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                string RutaArchivoMoverMI15 = pathString + @"\" + "MI15Final_" + ContadorArchivos.ToString() + ".csv"; ;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);
                System.IO.File.Move(RutaAmazon, RutaArchivoMoverMI15);
                ContadorArchivos += 1;

            }

            //textBox1.Text = "Fin Carga Archivos MI15";
            this.Refresh();
            this.Invalidate();
        }

        // Realiza accion de boton unifica reportes
        // ----------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Conexion = new SqlConnection();
                Conexion.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];

                SqlConnection ConexionGenerico1 = new SqlConnection();
                ConexionGenerico1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlGenerico1 = "ShippingByMarket";
                SqlCommand commandGenerico1 = new SqlCommand(sqlGenerico1, ConexionGenerico1);
                commandGenerico1.CommandType = CommandType.StoredProcedure;
                commandGenerico1.CommandTimeout = 7200; //in seconds
                ConexionGenerico1.Open();
                commandGenerico1.ExecuteNonQuery();
                ConexionGenerico1.Close();
            }
            catch (SystemException exp)
            {
                MessageBox.Show("Error: " + exp.Message);
                

            }
        }

   
        private void MainForm_Load(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
 
        }

         private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Increment(10);
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
            ProgressBar Progebar = new ProgressBar();
            

        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //int sum = 0;

            for(int ii=1; ii<=100; ii++)
            {

                //Thread.Sleep(100);
                //sum = sum + 1;
                //backgroundWorker1.ReportProgress(ii);
                //
                //if (backgroundWorker1.CancellationPending)
                //{
                //
                //    e.Cancel = true;
                //    backgroundWorker1.ReportProgress(0);
                //}

                //textBox1.Text = "Cantidad Registros :" + ContadorProgreso;
                this.Refresh();
                this.Invalidate();
            }
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void CargaCancelados(SqlConnection Conexion)
        {
            ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
            //clsInsertaRegistro.EliminaRegistroCancelados(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------

            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios Cancelados " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }
                FlgSihayFedex = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    ShippingByMarket.Clases.Cancelados clsPedido = new ShippingByMarket.Clases.Cancelados();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroCancelados(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDCancelados(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos Fedex";
            this.Refresh();
            this.Invalidate();
        }

        // Realiza accion de boton unifica reportes
        // ----------------------------------------
        private void EjecutaProceso()
        {
            try
            {
                SqlConnection Conexion = new SqlConnection();
                Conexion.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                ShippingByMarketMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ShippingByMarketMaxwarehouse.Clases.ManejoBD();
                ShippingByMarketMaxwarehouse.Clases.Logguer clsLogguer = new ShippingByMarketMaxwarehouse.Clases.Logguer();
                DateTime now = DateTime.Now;
                pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
                ReporteLog = "ReporteLog" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".txt";
                string pathOutPut1 = "";

                System.IO.Directory.CreateDirectory(pathString);

                //textBox1.Text = "Inicio Proceso";
                this.Refresh();
                this.Invalidate();

                now = DateTime.Now;
                clsLogguer.LogDuration(now, "Finaliza Insercion de BD DW...");
                //textBox1.Text = "Finaliza Insercion de BD DW...";

                // Inicia Carga de Fedex
                // ---------------------    
                ObtieneDatosFedex(Conexion);

                // Inicia carga de usps
                // --------------------
                ObtieneDatosUSPS(Conexion);

                // Inicia carga de UPS
                // -------------------
                ObtieneDatosUPS(Conexion);

                // Inicia carga Amazon
                // -------------------
                ObtieneDatosAmazon(Conexion);

                // Inicia carga Endicia
                // -------------------
                //textBox1.Text = "COMIENZA Carga Archivos Endicia";
                this.Refresh();
                this.Invalidate();

                ObtieneDatosEndicia(Conexion);

                //textBox1.Text = "Fin Carga Archivos Endicia";
                this.Refresh();
                this.Invalidate();

                //textBox1.Text = "Genera Excel";
                this.Refresh();
                this.Invalidate();
                using (var workbook = new XLWorkbook())
                {
                    // ejecuto sp que devuelve el crokis
                    // ----------------------------------
                    SqlConnection Conexion1 = new SqlConnection();

                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"]; 

                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("GeneraReporte", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader = cmd.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt.Load(reader);

                    var worksheet = workbook.Worksheets.Add("Output");

                    List<string[]> titles = new List<string[]> { new string[] { "SalesOrderNumber", "HoldCode", "TotalSales", "SalesSku", "SalesCategoryAtTimeOfSale", "UomCode", "UomQuantity", "SalesStatus", "SalesOrderDate", "SalesChannelName", "CustomerName", "FulfillmentSku", "FulfillmentChannelName", "FulfillmentChannelType", "LinkedFulfillmentChannelName", "FulfillmentLocationName", "FulfillmentOrderNumber", "Quantity", "Sku", "Title", "TotalCost", "Commission", "InventoryCost", "UnitCost", "ServiceCost", "EstimatedShippingCost", "ShippingCost", "ShippingPrice", "OverheadCost", "PackageCost", "ProfitLoss", "Carrier", "ShippingServiceLevel", "ShippedByUser", /*"ShippingWeight", "Length", "Width", "Height", "Weight", */"StateRegion", "TrackingNum", "MfrName", "PricingRule", "RequestedServiceLevel", "GroundTrackingIDPrefix", "ExpressorGroundTrackingID", "NetChargeAmount", "ServiceType", "GroundService", "ShipmentDate", "PODDeliveryDate", "ActualWeightAmount", "RatedWeightAmount", "DimLength", "DimWidth", "DimHeight", "DimDivisor", "ShipperState", "ZoneCode", "TenderedDate", "EarnedDiscount", "FuelSurcharge", "PerformancePricing", "DeliveryAreaSurchargeExtended", "DeliveryAreaSurcharge", "USPSNonMachSurcharge", "Residential", "GraceDiscount", "DeclaredValue", "DASExtendedResi", "AdditionalHandling", "ParcelReLabelCharge", "IndirectSignature", "DASResi", "AddressCorrection", "DASExtendedComm", "OversizeCharge", "AHSDimensions", "InvoiceDate", "InvoiceNumber", "MailClass", "TrackingNumberUSPS", "TotalPostageAmt", "DeliveryDate", "WeightUSPS", "Zone", "ServiceTypeUPS", "TrackingNumberUPS", "NetChargeAmountUPS", "order_idAmazon", "CarrierCargado" } };

                    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row

                    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathOutPut = pathOutPut + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                    workbook.SaveAs(pathOutPut);

                }

                //textBox1.Text = "Ejecuta SP No facturadas";
                this.Refresh();
                this.Invalidate();

                // Genera promedios FEDEX
                // ----------------------
                SqlConnection ConexionGenerico = new SqlConnection();
                ConexionGenerico.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlGenerico = "GeneraReporteNoFacturadas";
                SqlCommand commandGenerico = new SqlCommand(sqlGenerico, ConexionGenerico);
                commandGenerico.CommandType = CommandType.StoredProcedure;
                commandGenerico.CommandTimeout = 7200; //in seconds
                ConexionGenerico.Open();
                commandGenerico.ExecuteNonQuery();
                ConexionGenerico.Close();

                // Genera Reporte no facturadas
                // ----------------------------
                string PSNoFactutada = ConfigurationManager.AppSettings["RutaPSNoFacturada"];
                var proc7 = new System.Diagnostics.ProcessStartInfo();
                //string anyCommand;
                proc7.UseShellExecute = true;
                proc7.WorkingDirectory = @"C:\Windows\System32";
                proc7.FileName = @"C:\Windows\System32\cmd.exe";
                //proc1.Verb = "runas";
                proc7.Arguments = "/c " + "powershell -ExecutionPolicy Bypass -File \"" + PSNoFactutada + "\"";
                proc7.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                System.Diagnostics.Process.Start(proc7);

                // llena tabla de bianalitics
                // --------------------------
                //textBox1.Text = "Llena BI analitics";
                this.Refresh();
                this.Invalidate();
                SqlConnection Conexionbianalitics = new SqlConnection();
                Conexionbianalitics.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlbianalitics = "LlenaBIANALITICS";
                SqlCommand commandbianalitics = new SqlCommand(sqlbianalitics, Conexionbianalitics);
                commandbianalitics.CommandType = CommandType.StoredProcedure;
                commandbianalitics.CommandTimeout = 7200; //in seconds
                Conexionbianalitics.Open();
                commandbianalitics.ExecuteNonQuery();
                Conexionbianalitics.Close();

                // traslada facturas a historicos
                // ------------------------------
                //textBox1.Text = "Traslada historicos";
                this.Refresh();
                this.Invalidate();
                SqlConnection Conexionhistcarrier = new SqlConnection();
                Conexionhistcarrier.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlhistcarrier = "TrasladaCarrierHistorico";
                SqlCommand commandhistcarrier = new SqlCommand(sqlhistcarrier, Conexionhistcarrier);
                commandhistcarrier.CommandType = CommandType.StoredProcedure;
                commandhistcarrier.CommandTimeout = 7200; //in seconds
                Conexionhistcarrier.Open();
                commandhistcarrier.ExecuteNonQuery();
                Conexionhistcarrier.Close();

//                textBox1.Text = "Reporte Shipware";
                this.Refresh();
                this.Invalidate();
                using (var workbook = new XLWorkbook())
                {
                    // ejecuto sp que devuelve el crokis
                    // ----------------------------------
                    SqlConnection Conexion1 = new SqlConnection();

                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];

                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("ReporteShipware", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader = cmd.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt.Load(reader);

                    var worksheet = workbook.Worksheets.Add("Output");

                    List<string[]> titles = new List<string[]> { new string[] { "TRACKING_NUMBER", "Length", "Width", "Height", "Weight", "DimLength", "DimWidth", "DimHeight", "ActualWeightAmount", "AdditionalHandling", "OversizeCharge", "USPSNonMachSurcharge", "AHSDimensions" } };

                    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row

                    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathOutPut = pathOutPut + @"\ReporteShipware" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                    workbook.SaveAs(pathOutPut);

                }

                // Genera no facturadas FEDEX
                // ----------------------
                SqlConnection ConexionGenerico1 = new SqlConnection();
                ConexionGenerico1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlGenerico1 = "GeneraReporteNoFacturadas";
                SqlCommand commandGenerico1 = new SqlCommand(sqlGenerico1, ConexionGenerico1);
                commandGenerico1.CommandType = CommandType.StoredProcedure;
                commandGenerico1.CommandTimeout = 7200; //in seconds
                ConexionGenerico1.Open();
                commandGenerico1.ExecuteNonQuery();
                ConexionGenerico1.Close();

                //textBox1.Text = "Reporte Provision de Shipping";
                //this.Refresh();
                //this.Invalidate();
                //using (var workbook = new XLWorkbook())
                //{
                //    // ejecuto sp que devuelve el crokis
                //    // ----------------------------------
                //    SqlConnection Conexion1 = new SqlConnection();
                //
                //    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                //
                //    Conexion1.Open();
                //    SqlCommand cmd = new SqlCommand("ProvisionShipping", Conexion1);
                //    cmd.CommandType = CommandType.StoredProcedure;
                //    cmd.CommandTimeout = 7200; //in seconds
                //    SqlDataReader reader = cmd.ExecuteReader();
                //
                //    //Create a new DataTable.
                //    System.Data.DataTable dt = new System.Data.DataTable("Resultado");
                //
                //    //Load DataReader into the DataTable.
                //    dt.Load(reader);
                //
                //    var worksheet = workbook.Worksheets.Add("Output");
                //
                //    List<string[]> titles = new List<string[]> { new string[] { "TRACKING_NUMBER", "Length", "Width", "Height", "Weight", "DimLength", "DimWidth", "DimHeight", "ActualWeightAmount", "AdditionalHandling", "OversizeCharge", "USPSNonMachSurcharge", "AHSDimensions" } };
                //
                //    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row
                //
                //    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                //    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                //    pathOutPut = pathOutPut + @"\ReporteShipware" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                //    workbook.SaveAs(pathOutPut);
                //
                //}

                //textBox1.Text = "Inicio Proceso";
                this.Refresh();
                this.Invalidate();

                now = DateTime.Now;
                clsLogguer.LogDuration(now, "Genera dimensiones Promedio...");
               // textBox1.Text = "Genera dimensiones Promedio";

                using (var workbook = new XLWorkbook())
                {
                    // ejecuto sp que devuelve el crokis
                    // ----------------------------------
                    SqlConnection Conexion1 = new SqlConnection();

                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];

                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("DimensionesCompleto", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 1200; //in seconds
                    SqlDataReader reader = cmd.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt.Load(reader);

                    var worksheet = workbook.Worksheets.Add("Output");

                    List<string[]> titles = new List<string[]> { new string[] { "Sku", "UomCode", "ActualWeightAmount", "DimLength", "DimWidth", "DimHeight" } };

                    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row

                    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathOutPut = pathOutPut + @"\ShippingByMarket" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                    workbook.SaveAs(pathOutPut);
                }

                string ArchivoOrigenFTP = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                ArchivoOrigenFTP = ArchivoOrigenFTP + @"\ShippingByMarket" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                string ArchivoFinalFTP = ConfigurationManager.AppSettings["DireccionFTPDimensiones"];
                ArchivoFinalFTP = ArchivoFinalFTP + @"\ShippingByMarket" + DateTime.Now.ToString("MMddyyyy") + ".xlsx        ";

                SubeFTP(ArchivoOrigenFTP, ArchivoFinalFTP, true);


                // mueve archivo a carpeta de vera
                // -------------------------------
               // textBox1.Text = "Mueve Archivo Vera";
                this.Refresh();
                this.Invalidate();
                System.Threading.Thread.Sleep(600000);
                string PathVera = ConfigurationManager.AppSettings["RutaArchivosOutputsVera"];
                PathVera = PathVera + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                string PathDX = ConfigurationManager.AppSettings["RutaArchivosOutputsDX"];
                PathDX = PathDX + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";

                pathOutPut1 = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                pathOutPut1 = pathOutPut1 + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";

                System.IO.File.Copy(pathOutPut1, PathVera);
                System.IO.File.Copy(pathOutPut1, PathDX);

                System.Threading.Thread.Sleep(600000);
                //textBox1.Text = "MueveArchivoDX";
                this.Refresh();
                this.Invalidate();
                PathVera = ConfigurationManager.AppSettings["RutaArchivosOutputsVera"];
                PathVera = PathVera + @"/ReporteNoFacturadas" + DateTime.Now.ToString("MMddyyyy") + ".csv";
                PathDX = ConfigurationManager.AppSettings["RutaArchivosOutputsDX"];
                PathDX = PathDX + @"/ReporteNoFacturadas" + DateTime.Now.ToString("MMddyyyy") + ".csv";

                pathOutPut1 = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                pathOutPut1 = pathOutPut1 + @"\NoFacturadas" + DateTime.Now.ToString("MMddyyyy") + ".csv";

                System.IO.File.Copy(pathOutPut1, PathVera);
                System.IO.File.Copy(pathOutPut1, PathDX);

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error: " + exp.Message);
            }
        }

        public void SubeFTP(string ArchivoOrigenFTP, string ArchivoFinalFTP, bool EsExcel =false)
        {
            string UserFTP = ConfigurationManager.AppSettings["UserFTP"];
            string PassFTP = ConfigurationManager.AppSettings["PassFTP"];

            if (System.IO.File.Exists(ArchivoOrigenFTP))

            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ArchivoFinalFTP);
                request.Method = WebRequestMethods.Ftp.UploadFile;

                // This example assumes the FTP site uses anonymous logon.
                request.Credentials = new NetworkCredential(UserFTP, PassFTP);

                // finaliza conexion FTP
                // ---------------------
                request.KeepAlive = false;

                byte[] fileContents;
                // Si no es Excel utiliza este formateo de datos
                // ---------------------------------------------
                if (EsExcel == false)
                {
                    // Copy the contents of the file to the request stream.

                    using (StreamReader sourceStream = new StreamReader(ArchivoOrigenFTP))
                    {
                        fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                    }

                    request.ContentLength = fileContents.Length;
                }
                else 
                {
                    // Copy the contents of the file to the request stream.
                    fileContents = File.ReadAllBytes(ArchivoOrigenFTP);

                    request.ContentLength = fileContents.Length;
                }


                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(fileContents, 0, fileContents.Length);
                }

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    //Console.WriteLine($"Upload File Complete, status {response.StatusDescription}");

                    //MessageBox.Show("Subieron los archivos " + response.StatusDescription);
                    DateTime now = DateTime.Now;
                    ShippingByMarketMaxwarehouse.Clases.Logguer clsLogguer = new ShippingByMarketMaxwarehouse.Clases.Logguer();
                    clsLogguer.LogDuration(now, "Finaliza Carga de archivo"+ ArchivoOrigenFTP + " a direccion FTP " + ArchivoFinalFTP);
                    //textBox1.Text = "Finaliza Carga de archivo" + ArchivoOrigenFTP;
                }

            }

        }
        private void AsignaFechaInicioFin(DateTime Fecha1, DateTime Fecha2)
        {
 
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    // obtengo datos facturas
                    // ----------------------
                    SqlConnection Conexion1 = new SqlConnection();
                
                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"]; 
                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("DatosShippingByMarket", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 7200; //in seconds
                    cmd.Parameters.Add("@Operacion", SqlDbType.Int).Value = 5;
                    SqlDataReader reader = cmd.ExecuteReader();
                
                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");
                
                    //Load DataReader into the DataTable.
                    dt.Load(reader);
                    var worksheet = workbook.Worksheets.Add("shippingbymarket");

                    if (dt.Rows.Count != 0)
                    {
                        worksheet.Cell(1, 1).InsertData(dt);// inserta Contenido
                    }

                    string pathUPS = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathUPS = pathUPS + @"\ShippingByMarket" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(pathUPS);
                }
            }
            catch (SystemException exp)
            {
                MessageBox.Show("Error: " + exp.Message);


            }
        }

        private void label6_Click(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
