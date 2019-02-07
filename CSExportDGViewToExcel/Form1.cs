/****************************** Module Header ******************************\
* Module Name:  Form1.cs
* Project:      CSExportDGViewToExcel
* Copyright (c) Microsoft Corporation.
*
* The sample demostrates how to export DataGridView to Excel.
*
*
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/en-us/openness/resources/licenses.aspx#MPL.
* All other rights reserved.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/


using System;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace CSExportDGViewToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateRows();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        /// <summary>
        /// Add dummy rows to the datagridview.
        /// </summary>
        private void PopulateRows()
        {
            for (int i = 1; i <= 10; i++)
            {
                DataGridViewRow row =
                    (DataGridViewRow)dgvCityDetails.RowTemplate.Clone();

                row.CreateCells(dgvCityDetails, string.Format("City{0}", i),
                    string.Format("State{0}", i), string.Format("Country{0}", i));

                dgvCityDetails.Rows.Add(row);

            }
        }

        /// <summary>
        /// Exports the datagridview values to Excel.
        /// </summary>
        private void ExportToExcel()
        {

            XLWorkbook Workbook = new XLWorkbook("C:\\ERP_Monitoring.xlsx");
            var worksheet = Workbook.Worksheet("ERPREPORT");


            try
            {
                //worksheet.Name = "ExportedFromDatGrid";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column.
                for (int i = 0; i < dgvCityDetails.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dgvCityDetails.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cell(cellRowIndex, cellColumnIndex).Value = dgvCityDetails.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cell(cellRowIndex, cellColumnIndex).Value = dgvCityDetails.Rows[i].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                Workbook.Save();

                //Getting the location and file name of the excel to save from user.
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;                

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Workbook.SaveAs(saveDialog.FileName + ".xlsx");
                    MessageBox.Show("Export Successful");
                }               
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Diagnostics.Process ps = new System.Diagnostics.Process();
                ps.StartInfo.FileName = "C:\\ERP_Monitoring.xlsx";
                ps.Start();
                ps.Dispose();
            }

        }
    }
}
