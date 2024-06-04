# VOI_v.3.1.2

            Excel.Workbook wb = excapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet sheet = (Excel.Worksheet)wb.ActiveSheet;
            var wb = excapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var sheet = (Excel.Worksheet)wb.ActiveSheet;
            MakeCaption(1, sheet, reportData);
            object[,] data = new object[reportData.Rows.Count, reportData.Columns.Count];
            var data = new object[reportData.Rows.Count, reportData.Columns.Count];
            for (int i = 0; i < reportData.Rows.Count; i++)
            for (var i = 0; i < reportData.Rows.Count; i++)
            {
                for (int j = 0; j < reportData.Columns.Count; j++)
                for (var j = 0; j < reportData.Columns.Count; j++)
                {
                    data[i, j] = reportData.Rows[i][j];
                }
            }
            Excel.Range rg = sheet.Range[sheet.Cells[2, 1], sheet.Cells[1 + reportData.Rows.Count, reportData.Columns.Count]];
            var rg = sheet.Range[sheet.Cells[2, 1], sheet.Cells[1 + reportData.Rows.Count, reportData.Columns.Count]];
            rg.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, data);
            SetNumberFormat(1, sheet, reportData, reportData.Rows.Count);
