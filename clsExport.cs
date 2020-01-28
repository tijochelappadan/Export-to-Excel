using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reports
{
    public class clsExport
    {
        public void ExportToExcel(DataGridView dgv, string ReportName, string ReportHeader, string Head1 = "", string Head2 = "", string Head3 = "")
        {
            DataGridView dgvDetails = CopyDataGridView(dgv);

            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = ReportName;
            //// Inserting Company Details
            string strColumnAlphabet = GetColumnAlphabet(dgvDetails.Columns.Count); // ((char)(dgvDetails.Columns.Count + 64)).ToStringCustom();            
            worksheet.Cells[1, 1] = ClsCommonSettings.GlobalCompany;
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range["A1:" + strColumnAlphabet + "1"].Merge();

            worksheet.Cells[2, 1] = "GSTIN : " + ClsCommonSettings.CompanyTINNo + ", " + ClsCommonSettings.CompanyAddress;
            worksheet.Cells[2, 1].Font.Bold = true;
            worksheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range["A2:" + strColumnAlphabet + "2"].Merge();

            worksheet.Cells[3, 1] = "";
            worksheet.Range["A3:" + strColumnAlphabet + "3"].Merge();

            worksheet.Cells[4, 1] = ReportName;
            worksheet.Cells[4, 1].Font.Bold = true;
            worksheet.Cells[4, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range["A4:" + strColumnAlphabet + "4"].Merge();

            worksheet.Cells[5, 1] = ReportHeader;
            worksheet.Cells[5, 1].Font.Bold = true;
            worksheet.Range["A5:" + strColumnAlphabet + "5"].Merge();

            worksheet.Cells[6, 1] = "";
            worksheet.Range["A6:" + strColumnAlphabet + "6"].Merge();

            int RowIndex = 7;

            if (Head1 != string.Empty)
            {
                worksheet.Cells[RowIndex, 1] = Head1;
                worksheet.Cells[RowIndex, 1].Font.Bold = true;
                worksheet.Range["A" + RowIndex + ":" + strColumnAlphabet + RowIndex.ToStringCustom()].Merge();
                RowIndex = RowIndex + 1;
            }

            if (Head2 != string.Empty)
            {
                worksheet.Cells[RowIndex, 1] = Head2;
                worksheet.Cells[RowIndex, 1].Font.Bold = true;
                worksheet.Range["A" + RowIndex + ":" + strColumnAlphabet + RowIndex.ToStringCustom()].Merge();
                RowIndex = RowIndex + 1;
            }

            if (Head3 != string.Empty)
            {
                worksheet.Cells[RowIndex, 1] = Head3;
                worksheet.Cells[RowIndex, 1].Font.Bold = true;
                worksheet.Range["A" + RowIndex + ":" + strColumnAlphabet + RowIndex.ToStringCustom()].Merge();
                RowIndex = RowIndex + 1;
            }

            if (RowIndex > 7)
                RowIndex = RowIndex + 1;
            
            // storing header part in Excel  
            for (int i = 1; i < dgvDetails.Columns.Count + 1; i++)
            {
                worksheet.Cells[RowIndex, i] = dgvDetails.Columns[i - 1].HeaderText;
                worksheet.Cells[RowIndex, i].Font.Bold = true;
            }

            RowIndex = RowIndex + 2;
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dgvDetails.Rows.Count; i++)
            {
                for (int j = 0; j < dgvDetails.Columns.Count; j++)
                {
                    if (dgvDetails.Columns[j].Name.Contains("BillNo") || dgvDetails.Columns[j].Name.Contains("GSTIN") || dgvDetails.Columns[j].Name.Contains("InvoiceNo") ||
                        dgvDetails.Columns[j].Name.Contains("HSN") || dgvDetails.Columns[j].Name.Contains("Description") || dgvDetails.Columns[j].Name.Contains("TinNo") ||
                        dgvDetails.Columns[j].Name.Contains("SlNo"))
                    {
                        worksheet.Cells[i + RowIndex, j + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                        worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "@";
                        worksheet.Cells[i + RowIndex, j + 1] = dgvDetails.Rows[i].Cells[j].Value.ToStringCustom();
                    }
                    else if (dgvDetails.Columns[j].Name.Contains("Date") && dgvDetails.Rows[i].Cells[j].Value.ToStringCustom() != string.Empty)
                    {
                        worksheet.Cells[i + RowIndex, j + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                        worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "@";
                        worksheet.Cells[i + RowIndex, j + 1] = ReportName == "B2B" ? dgvDetails.Rows[i].Cells[j].Value.ToStringCustom() : dgvDetails.Rows[i].Cells[j].Value.ToDateTime().ToShortDateString();
                    }
                    else if (dgvDetails.Columns[j].Name == "Balance")
                    {
                        worksheet.Cells[i + RowIndex, j + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                        worksheet.Cells[i + RowIndex, j + 1] = dgvDetails.Rows[i].Cells[j].Value.ToStringCustom();
                    }
                    else
                    {
                        decimal value;

                        if (Decimal.TryParse(dgvDetails.Rows[i].Cells[j].Value.ToStringCustom(), out value))
                        {
                            if (dgvDetails.Columns[j].Name.Contains("Qty") || dgvDetails.Columns[j].Name.Contains("Quantity") || 
                                dgvDetails.Columns[j].Name.Contains("Stock"))
                            {
                                if (value == 0)
                                    worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "0.000";
                                else
                                    worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "###################.000";
                            }
                            else
                            {
                                switch (ClsCommonSettings.DecimalPlaces)
                                {
                                    case "1":
                                        if (value == 0)
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "0.0";
                                        else
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "###################.0";
                                            break;
                                    case "2":
                                        if (value == 0)
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "0.00";
                                        else
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "###################.00";
                                        break;
                                    case "3":
                                        if (value == 0)
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "0.000";
                                        else
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "###################.000";
                                        break;
                                    case "4":
                                        if (value == 0)
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "0.0000";
                                        else
                                            worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "###################.0000";
                                        break;
                                }
                            }

                            worksheet.Cells[i + RowIndex, j + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                        }
                        else if (dgvDetails.Columns[j].Name == "ExpenseAmt" || dgvDetails.Columns[j].Name == "IncomeAmt" ||
                            dgvDetails.Columns[j].Name == "LiabilityAmt" || dgvDetails.Columns[j].Name == "AssetAmt")
                        {
                            if (dgvDetails.Rows[i].Cells[j].Value.ToStringCustom().Contains("("))
                            {
                                worksheet.Cells[i + RowIndex, j + 1].NumberFormat = "@";
                                worksheet.Cells[i + RowIndex, j + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                            }
                        }

                        worksheet.Cells[i + RowIndex, j + 1] = dgvDetails.Rows[i].Cells[j].Value.ToStringCustom();
                    }
                }
            }
            // Auto fit columns
            worksheet.Columns.AutoFit();
            worksheet.PageSetup.TopMargin = 0.75;
            worksheet.PageSetup.LeftMargin = 0.25;
            worksheet.PageSetup.RightMargin = 0.25;
            worksheet.PageSetup.BottomMargin = 0.75;
            worksheet.PageSetup.CenterHorizontally = true;
            worksheet.PageSetup.PrintGridlines = true;
            // Exit from the application  
            app.Quit();  
        }

        private string GetColumnAlphabet(int ColumnCount)
        {
            if (ColumnCount > 0 && ColumnCount < 27)
                return ((char)(ColumnCount + 64)).ToStringCustom();
            else
                return "A" + ((char)((ColumnCount - 26) + 64)).ToStringCustom();
        }

        private DataGridView CopyDataGridView(DataGridView dgv)
        {
            DataGridView dgv_copy = new DataGridView();

            try
            {
                if (dgv_copy.Columns.Count == 0)
                {
                    foreach (DataGridViewColumn dgvc in dgv.Columns)
                    {
                        if (dgvc.Visible)
                            dgv_copy.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                    }
                }

                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    int intColIndex = 0;
                    dgv_copy.Rows.Add();

                    foreach (DataGridViewCell cell in dgv.Rows[i].Cells)
                    {
                        if (cell.Visible)
                        {
                            dgv_copy.Rows[i].Cells[intColIndex].Value = cell.Value;
                            intColIndex++;
                        }
                    }
                }

                dgv_copy.AllowUserToAddRows = false;
                dgv_copy.Refresh();

            }
            catch { }

            return dgv_copy;
        }
    }
}
