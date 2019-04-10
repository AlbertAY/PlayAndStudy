using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool
{
    public class SpreadJsCreate
    {


        public static string GetJson()
        {
            int rowCount = 5000, columnCount = 15;
            List<List<Cell>> rowsData = new List<List<Cell>>();

            List<row> rows = new List<row>();
            DataTable dataTable = new DataTable();

            for (int i = 0; i < columnCount; i++)
            {
                dataTable.Columns.Add(i.ToString(), typeof(Cell));
            }

            

            for (int i = 0; i < rowCount; i++)
            {
                DataRow dataRow = dataTable.NewRow();

                List<Cell> rowCells = new List<Cell>();
                for (int j = 0; j < columnCount; j++)
                {
                    dataRow[j.ToString()] = new Cell();

                    Cell cell = new Cell();

                    rowCells.Add(cell);
                }
                //dataRow.ItemArray = rowCells.ToArray();
                //rowsData.Add(rowCells);
                rows.Add(new row());

                dataTable.Rows.Add(dataRow);

            }
            List<column> columns = new List<column>();
            for (int i = 0; i < columnCount; i++)
            {
                columns.Add(new column());
            }



            dynamic data = new Sheet();
            data.name = "性能测试";
            data.rowCount = rowCount;
            data.columnCount = columnCount;
            data.frozenColCount = 1;
            data.frozenRowCount = 2;
            data.theme = "Office";
            data.data = new ExeclData(dataTable);
            data.rows = rows;
            data.columns = columns;

            WrokBook wrokBook = new WrokBook();
            sheetData sheetData = new sheetData();
            sheetData.测试数据 = data;

            wrokBook.sheets = sheetData;


            string jsonData = JsonConvert.SerializeObject(wrokBook);
            return jsonData;


        }
    }


    public class Cell
    {
        public string value => "单元格数据";

        public style style => new style();
    }

    public class style
    {
        borderstyle borderstyle = null;
        public style()
        {
            borderstyle = new borderstyle();

        }
        public string backColor => "#F4F4F9";

        public int hAlign => 1;

        public int vAlign => 1;

        public string font => "14.6667px Calibri";

        public string themeFont => "Body";

        public borderstyle borderLeft => borderstyle;

        public borderstyle borderTop => borderstyle;

        public borderstyle borderRight => borderstyle;

        public borderstyle borderBottom => borderstyle;

        public bool locked => true;

        public int imeMode => 1;

    }

    public class borderstyle
    {
        public int style => 1;
    }


    public class Sheet
    {
        public string name { set; get; }
        public int rowCount { set; get; }
        public int columnCount { set; get; }
        public int frozenColCount { set; get; }
        public int frozenRowCount { set; get; }
        public string theme { set; get; }
        public ExeclData data { set; get; }

        public List<column> columns { set; get; }

        public List<row> rows { set; get; }

        public protectionOption protectionOptions => new protectionOption();

        public outlineColumnOptions outlineColumnOptions = new outlineColumnOptions();

        public bool showRowOutline = false;

        public List<spans> spans => null;

    }

    public class ExeclData
    {
        public ExeclData(DataTable data)
        {
            dataTable = data;
        }
        public DataTable dataTable { set; get; }
    }

    public class WrokBook
    {
        public string version => "11.2.2";

        public sheetData sheets { set; get; }

    }

    public class sheetData
    {
        public Sheet 测试数据 { set; get; }
    }


    public class column
    {
        public int size => 200;
    }

    public class row
    {
        public int size => 20;
    }

    public class protectionOption
    {
        public bool allowSelectLockedCells => true;
        public bool allowSelectUnlockedCells => true;
        public bool allowCopyPasteExcelStyle => true;
        public bool allowFilter => true;
        public bool allowSort => true;
        public bool allowResizeRows => true;
        public bool allowResizeColumns => true;
        public bool allowEditObjects => true;
        public bool autoFitType => true;

        public bool tabStripVisible => false;
        public bool tabEditable => false;
        public bool newTabVisible => false;
        public bool allowDragInsertRows => false;
        public bool allowDragInsertColumns => false;
        public bool allowInsertRows => false;
        public bool allowInsertColumns => false;
        public bool allowDeleteRows => false;

        public bool allowDeleteColumns => false;
    }


    public class outlineColumnOptions
    {
        public int columnIndex => 0;
        public bool showCheckBox => false;
        public int maxLevel => 2;
    }

    public class spans
    {
        public int row { set; get; }
        public int rowCount { set; get; }
        public int col { set; get; }
        public int colCount { set; get; }
    }

}
