using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using xls=Microsoft.Office.Interop.Excel;

namespace xcel_viewr
{
    public partial class ExcelForm : System.Web.UI.Page
    {
        private xls.Application appOP = null;
        protected static string m_strFileName = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (appOP == null)
            {
                appOP = new xls.Application();
            }
            txtfileValue.EnableViewState = true;
        }

        protected override void OnUnload(EventArgs e)
        {
            try
            {
                if (appOP != null)
                {
                    appOP.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(appOP);
                    appOP = null;
                }
            }
            catch (Exception eqq)
            {
                Response.Write(eqq.ToString());
            }
            base.OnUnload(e);
        }

        protected void btnAvailableShtAndChrt_Click(object sender, EventArgs e)
        {
            m_strFileName = Server.MapPath("/");
            m_strFileName += @"\Uploads\uploadedfile.xls";
            Response.Write("File name:: " + m_strFileName);

            if (m_strFileName == "")
            {
                lblErrText.Text = "File is not Available";
            }
            else
            {
                Response.Write("File Found");
                string strTemp = m_strFileName.Substring(m_strFileName.Length - 3);
                strTemp = strTemp.ToUpper();
                drpShtAndChrt.Items.Clear();
                GetListofSheetsAndCharts(m_strFileName, true, drpShtAndChrt);

                //if (strTemp == "XLS" || strTemp == "XLSM")
                //{
                //    drpShtAndChrt.Items.Clear();
                //    GetListofSheetsAndCharts(m_strFileName, true, drpShtAndChrt);
                //}
                //else
                //{
                //    lblErrText.Text = "Selected File is not Required Format";
                //}
            }
        }

        protected void btnDisplay_Click(object sender, EventArgs e)
        {

            if (drpShtAndChrt.SelectedIndex != -1)
            {
                string strSheetorChartName = drpShtAndChrt.SelectedItem.Text;
                // Because "*" cannot be accepted by Sheet Name in Excel
                char[] delimiterChars = { '*' };
                string[] strTemp = strSheetorChartName.Split(delimiterChars);

                if (strTemp[1] == "WorkSheet")
                {
                    Response.Write("Display Sheets");
                    DisplayExcelSheet(m_strFileName, strTemp[0], true, lblErrText);
                }
                else if (strTemp[1] == "Chart")
                {
                    Response.Write("Display Charts");
                    DisplayExcelSheet(m_strFileName, strTemp[0], true, lblErrText, true);
                }
            }

        }

        public void btnUpload_Click(object sender, EventArgs e)
        {
            string filepath = Server.MapPath("/");
            filepath += @"\Uploads\uploadedfile.xls";

            txtfileValue.PostedFile.SaveAs(filepath);
            Response.Write("File saved");
            return;
        }

        public void GetListofSheetsAndCharts(string strFileName, bool bReadOnly, DropDownList drpList)
        {
            Response.Write("Loading sheets");
            xls.Workbook workbook = null;
            try
            {
                Response.Write("<br>" + strFileName + "<br>");

                if (!bReadOnly)
                {
                    // Write Mode Open
                    workbook = appOP.Workbooks.Open(strFileName, 2, false, 5, "", "", true, xls.XlPlatform.xlWindows, "\t", false, true, 0, true, 1, 0);
                    // For Optimal Opening 
                    //workbook = appOP.Workbooks.Open(strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                else
                {
                    // Read Mode Open
                    workbook = appOP.Workbooks.Open(strFileName, 2, true, 5, "", "", true, xls.XlPlatform.xlWindows, "\t", false, true, 0, true, 1, 0);
                    // For Optimal Opening 
                    //workbook = appOP.Workbooks.Open(strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                // Reading of Excel File

                object SheetRChart = null;
                int nTotalWorkSheets = workbook.Sheets.Count;
                int nIndex = 0;
                //foreach (object SheetRChart in workbook.Sheets)

                for (int nWorkSheet = 1; nWorkSheet <= nTotalWorkSheets; nWorkSheet++)
                {
                    SheetRChart = workbook.Sheets[(object)nWorkSheet];
                    if (SheetRChart is xls.Worksheet)
                    {

                        ListItem lstItemAdd = new ListItem(((xls.Worksheet)SheetRChart).Name + "*WorkSheet", nIndex.ToString(), true);
                        drpList.Items.Add(lstItemAdd);
                        lstItemAdd = null;
                        nIndex++;
                    }
                    else if (SheetRChart is xls.Chart)
                    {
                        ListItem lstItemAdd = new ListItem(((xls.Chart)SheetRChart).Name + "*Chart", nIndex.ToString(), true);
                        drpList.Items.Add(lstItemAdd);
                        lstItemAdd = null;
                        nIndex++;
                    }
                }

                if (workbook != null)
                {
                    if (!bReadOnly)
                    {
                        // Write Mode Close
                        workbook.Save();
                        workbook = null;
                    }
                    else
                    {
                        // Read Mode Close
                        workbook.Close(false, false, Type.Missing);
                        workbook = null;
                    }
                }

            }
            catch (Exception expFile)
            {
                Response.Write(expFile.ToString());
            }
            finally
            {
                if (workbook != null)
                {
                    if (!bReadOnly)
                    {
                        // Write Mode Close
                        workbook.Save();
                        workbook = null;
                    }
                    else
                    {
                        // Read Mode Close
                        workbook.Close(false, false, Type.Missing);
                        workbook = null;
                    }
                }
            }
        }

        public bool DisplayExcelSheet(string strFileName, string strSheetRChartName, bool bReadOnly, Label lblErrorText)
        {
            return DisplayExcelSheet(strFileName, strSheetRChartName, bReadOnly, lblErrText, false);
        }
        /// <summary>
        /// Displaying a given Excel WorkSheet
        /// </summary>
        /// <param name="strFileName">The Filename to be selected</param>
        /// <param name="strSheetRChartName">The Sheet or Chart Name to be Displayed</param>
        /// <param name="bReadOnly">Specifies the File should be open in Read only mode,
        /// If it is true then the File will be open ReadOnly</param>
        /// <param name="lblErrorText">If any Error Occurs should be Displayed</param>
        /// <param name="bIsChart">Specifies whether it is a Chart</param>
        /// <returns>Returns Boolean Value the Method Succeded</returns>
        public bool DisplayExcelSheet(string strFileName, string strSheetRChartName, bool bReadOnly, Label lblErrorText, bool bIsChart)
        {

            appOP.DisplayAlerts = false;
            xls.Workbook workbook = null;
            xls.Worksheet worksheet = null;
            xls.Chart chart = null;

            try
            {

                if (!bReadOnly)
                {
                    // Write Mode Open
                    workbook = appOP.Workbooks.Open(strFileName, 2, false, 5, "", "", true, xls.XlPlatform.xlWindows, "\t", false, true, 0, true, 1, 0);
                    // For Optimal Opening 
                    //workbook = appOP.Workbooks.Open(strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                else
                {
                    // Read Mode Open
                    workbook = appOP.Workbooks.Open(strFileName, 2, true, 5, "", "", true, xls.XlPlatform.xlWindows, "\t", false, true, 0, true, 1, 0);
                    // For Optimal Opening 
                    //workbook = appOP.Workbooks.Open(strFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }



                // Reading of Excel File

                if (bIsChart)
                {
                    chart = (xls.Chart)workbook.Charts[strSheetRChartName];
                }
                else
                {
                    worksheet = (xls.Worksheet)workbook.Sheets[strSheetRChartName];
                }

                // Reading the File Information Codes goes Here
                if (bIsChart)
                {
                    if (chart == null)
                    {
                        lblErrorText.Text = strSheetRChartName + " Chart is Not Available";
                    }
                    else
                    {
                        ExcelChartRead(chart, this.pnlBottPane);
                    }
                }
                else
                {
                    if (worksheet == null)
                    {
                        lblErrorText.Text = strSheetRChartName + " Sheet is Available";
                    }
                    else
                    {
                        this.pnlBottPane.Controls.Add(ExcelSheetRead(worksheet, lblErrText));
                    }
                }

                if (!bReadOnly)
                {
                    // Write Mode Close
                    workbook.Save();
                    workbook = null;
                }
                else
                {
                    // Read Mode Close
                    workbook.Close(false, false, Type.Missing);
                    workbook = null;
                }
            }
            catch (Exception expInterop)
            {
                lblErrText.Text = expInterop.ToString();
                return false;
            }
            finally
            {
                if (workbook != null)
                {
                    if (!bReadOnly)
                    {
                        // Write Mode Close
                        workbook.Save();
                        workbook = null;
                    }
                    else
                    {
                        // Read Mode Close
                        workbook.Close(false, false, Type.Missing);
                        workbook = null;
                    }
                }
                appOP.DisplayAlerts = true;
            }
            return true;
        }

        /// <summary>
        /// To Display a Chart in the Panel Object
        /// </summary>
        /// <param name="objExcelChart">Chart to be Opened</param>
        /// <param name="ctrlCollPane">Panel Object to be Displayed</param>
        /// <returns>Returns Boolean Value the Method Succeded</returns>
        public bool ExcelChartRead(xls.Chart objExcelChart, Panel ctrlCollPane)
        {

            Image imgChart = null;
            try
            {
                objExcelChart.Export(Server.MapPath("/") + @"/tmp/TempGif.gif", "GIF", true);
                imgChart = new Image();
                imgChart.ImageUrl = @"/tmp/TempGif.gif";
                //Response.Write("<img src ='/tmp/TempGif.gif'/>");
                ctrlCollPane.Controls.Add(imgChart);
                ctrlCollPane.Visible = true;
                imgChart.Visible = true;
            }
            catch (Exception expFileError)
            {
                Response.Write(expFileError.ToString());
                return false;
            }
            //finally
            //{
            //    if (imgChart != null)
            //    {
            //        imgChart.Dispose();
            //    }
            //}
            return true;
        }


        /// <summary>
        /// Read an Excel Sheet and Displays as it is Same
        /// </summary>
        /// <param name="objExcelSheet">Worksheet to be displayed</param>
        /// <param name="lblErrText">If any Error Occurs that will be displayed</param>
        /// <returns>Returns a Table Control that contains Worksheet Information</returns>
        public Control ExcelSheetRead(xls.Worksheet objExcelSheet, Label lblErrText)
        {

            int nMaxCol = ((xls.Range)objExcelSheet.UsedRange).EntireColumn.Count;
            int nMaxRow = ((xls.Range)objExcelSheet.UsedRange).EntireRow.Count;


            Table tblOutput = new Table();

            TableRow TRow = null;
            TableCell TCell = null;

            string strSize = "";
            int nSizeVal = 0;
            bool bMergeCells = false;
            int nMergeCellCount = 0;
            int nWidth = 0;


            if (objExcelSheet == null)
            {
                return (Control)tblOutput;
            }

            tblOutput.CellPadding = 0;
            tblOutput.CellSpacing = 0;
            tblOutput.GridLines = GridLines.Both;


            try
            {

                for (int nRowIndex = 1; nRowIndex <= nMaxRow; nRowIndex++)
                {
                    TRow = null;
                    TRow = new TableRow();


                    for (int nColIndex = 1; nColIndex <= nMaxCol; nColIndex++)
                    {

                        TCell = null;
                        TCell = new TableCell();
                        if (((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Value2 != null)
                        {

                            TCell.Text = ((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Text.ToString();
                            if (((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Comment != null)
                            {
                                TCell.ForeColor = System.Drawing.Color.Blue;
                                TCell.ToolTip = ((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Comment.Shape.AlternativeText;
                            }
                            else
                            {
                                TCell.ForeColor = ConvertExcelColor2DotNetColor(((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Color);
                            }

                            TCell.BorderWidth = 2;
                            TCell.Width = 140; //TCell.Width = 40;

                            //*
                            TCell.Font.Bold = (bool)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Bold;
                            TCell.Font.Italic = (bool)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Italic;
                            strSize = ((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Size.ToString();
                            nSizeVal = Convert.ToInt32(strSize);
                            TCell.Font.Size = FontUnit.Point(nSizeVal);
                            TCell.BackColor = ConvertExcelColor2DotNetColor(((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Interior.Color);

                            if ((bool)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).MergeCells != false)
                            {
                                if (bMergeCells == false)
                                {
                                    TCell.ColumnSpan = (int)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).MergeArea.Columns.Count;
                                    nMergeCellCount = (int)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).MergeArea.Columns.Count;
                                    nMergeCellCount--;
                                    bMergeCells = true;
                                }
                                else if (nMergeCellCount == 0)
                                {
                                    TCell.ColumnSpan = (int)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).MergeArea.Columns.Count;
                                    nMergeCellCount = (int)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).MergeArea.Columns.Count;
                                    nMergeCellCount--;
                                }
                            }
                            else
                            {
                                bMergeCells = false;
                            }

                            TCell.HorizontalAlign = ExcelHAlign2DotNetHAlign(((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]));
                            TCell.VerticalAlign = ExcelVAlign2DotNetVAlign(((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]));
                            TCell.Height = Unit.Point(Decimal.ToInt32(Decimal.Parse((((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).RowHeight.ToString()))));
                            nWidth = Decimal.ToInt32(Decimal.Parse((((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).ColumnWidth.ToString())));
                            TCell.Width = Unit.Point(nWidth * nWidth);
                            //*/

                        }
                        else
                        {
                            if ((bool)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).MergeCells == false)
                            {
                                bMergeCells = false;
                            }
                            if (bMergeCells == true)
                            {
                                nMergeCellCount--;
                                continue;
                            }
                            TCell.Text = "&nbsp;";
                            if (((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Comment != null)
                            {
                                TCell.ForeColor = System.Drawing.Color.Blue;
                                TCell.ToolTip = ((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Comment.Shape.AlternativeText;
                            }
                            else
                            {
                                TCell.ForeColor = ConvertExcelColor2DotNetColor(((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Color);
                            }
                            TCell.Font.Bold = (bool)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Bold;
                            TCell.Font.Italic = (bool)((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Italic;
                            strSize = ((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Font.Size.ToString();
                            nSizeVal = Convert.ToInt32(strSize);
                            TCell.Font.Size = FontUnit.Point(nSizeVal);
                            TCell.BackColor = ConvertExcelColor2DotNetColor(((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).Interior.Color);

                            TCell.Height = Unit.Point(Decimal.ToInt32(Decimal.Parse((((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).RowHeight.ToString()))));
                            nWidth = Decimal.ToInt32(Decimal.Parse((((xls.Range)objExcelSheet.Cells[nRowIndex, nColIndex]).ColumnWidth.ToString())));
                            TCell.Width = Unit.Point(nWidth * nWidth);
                        }

                        //TCell.BorderStyle = BorderStyle.Solid;
                        //TCell.BorderWidth = Unit.Point(1);
                        //TCell.BorderColor = System.Drawing.Color.Gray;

                        TRow.Cells.Add(TCell);


                    }

                    tblOutput.Rows.Add(TRow);

                }
            }
            catch (Exception ex)
            {
                lblErrText.Text = ex.ToString();
            }
            return (Control)tblOutput;
        }

        /// <summary>
        /// Converts Excel Color to Dot Net Color
        /// </summary>
        /// <param name="objExcelColor">Excel Object Color</param>
        /// <returns>Returns System.Drawing.Color</returns>
        private System.Drawing.Color ConvertExcelColor2DotNetColor(object objExcelColor)
        {

            string strColor = "";
            uint uColor = 0;
            int nRed = 0;
            int nGreen = 0;
            int nBlue = 0;

            strColor = objExcelColor.ToString();
            uColor = checked((uint)Convert.ToUInt32(strColor));
            strColor = String.Format("{0:x2}", uColor);
            strColor = "000000" + strColor;
            strColor = strColor.Substring((strColor.Length - 6), 6);

            uColor = 0;
            uColor = Convert.ToUInt32(strColor.Substring(4, 2), 16);
            nRed = (int)uColor;

            uColor = 0;
            uColor = Convert.ToUInt32(strColor.Substring(2, 2), 16);
            nGreen = (int)uColor;

            uColor = 0;
            uColor = Convert.ToUInt32(strColor.Substring(0, 2), 16);
            nBlue = (int)uColor;

            return System.Drawing.Color.FromArgb(nRed, nGreen, nBlue);
        }


        /// <summary>
        /// Converts Excel Horizontal Alignment to DotNet Horizontal Alignment
        /// </summary>
        /// <param name="objExcelAlign">Excel Horizontal Alignment</param>
        /// <returns>HorizontalAlign</returns>
        private HorizontalAlign ExcelHAlign2DotNetHAlign(object objExcelAlign)
        {
            string exp = ((xls.Range)objExcelAlign).HorizontalAlignment.ToString();
            switch (exp)
            {
                case "-4131":
                    return HorizontalAlign.Left;
                case "-4108":
                    return HorizontalAlign.Center;
                case "-4152":
                    return HorizontalAlign.Right;
                default:
                    return HorizontalAlign.Left;
            }
        }

        /// <summary>
        /// Converts Excel Vertical Alignment to DotNet Vertical Alignment
        /// </summary>
        /// <param name="objExcelAlign">Excel Vertical Alignment</param>
        /// <returns>VerticalAlign</returns>
        private VerticalAlign ExcelVAlign2DotNetVAlign(object objExcelAlign)
        {
            string exp = ((xls.Range)objExcelAlign).VerticalAlignment.ToString();
            switch (exp)
            {
                case "-4160":
                    return VerticalAlign.Top;
                case "-4108":
                    return VerticalAlign.Middle;
                case "-4107":
                    return VerticalAlign.Bottom;
                default:
                    return VerticalAlign.Bottom;
            }

        }

    }
}