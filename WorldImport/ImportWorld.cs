using System;
using System.Collections.Generic;
using System.Text;
using Spire.Xls;
using Spire.Pdf;
using Spire.Doc;
using System.Data;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace WorldImport
{
    public class ImportWorld
    {
        public void Import(string path)
        {
            var doc = new Document();
            doc.LoadFromFile(path);
            doc.Replace("$CONTRACT", "2021NTI-ROBOT-P366-RBJLDDJ", true, true);
            doc.Replace("$A&B", "测试导出", true, true);
            doc.Replace("$A", "深圳市今天国际物流股份有限公司", true, true);
            doc.Replace("$B", "深圳市视酷信息技术有限公司", true, true);
            doc.Replace("$Y", "2021", true, true);
            doc.Replace("$M", "3", true, true);
            doc.Replace("$D", "2", true, true);
            DataTable dt = new DataTable();
            //1.创建空列
            DataColumn dc = new DataColumn();
            //2.创建带列名和类型名的列(两种方式任选其一)
            dt.Columns.Add("序号");
            dt.Columns.Add("名称");
            dt.Columns.Add("规格型号");
            dt.Columns.Add("品牌/厂家");
            dt.Columns.Add("数量");
            dt.Columns.Add("单位");
            dt.Columns.Add("单价");
            dt.Columns.Add("总价");
            dt.Columns.Add("备注");
            DataRow dr = dt.NewRow();
            dr["序号"] = 1;
            dr["名称"] = "电脑";
            dr["规格型号"] = "R720";
            dr["品牌/厂家"] ="联想";
            dr["数量"] = 10;
            dr["单位"] = "台";
            dr["单价"] = 8200;
            dr["总价"] = 82000;
            dr["备注"] = "测试";
            dt.Rows.Add(dr);
            for (int i = 0; i < 20; i++)
            {
                dr = dt.NewRow();
                dr["序号"] = 2+i;
                dr["名称"] = "手机"+i;
                dr["规格型号"] = "iPone8"+i;
                dr["品牌/厂家"] = "iPone"+i;
                dr["数量"] = 20+i;
                dr["单位"] = "个";
                dr["单价"] = 5000+i;
                dr["总价"] = 100000+i;
                dr["备注"] = 1+i;
                dt.Rows.Add(dr);
            }
            GetTable(dt,doc,60);
            string savePath = @"C:\Users\zengwang\Desktop\采购合同范本-修订稿-20210105\测4\123.docx"; //导出路径
            doc.SaveToFile(savePath, Spire.Doc.FileFormat.Docx);
            doc.Close();
        }
        public void GetTable(DataTable dt,Document doc,decimal disconuts)
        {
            Section section = doc.Sections[1];
            Paragraph p;
            Table table = section.AddTable(true);
            table.ResetCells(dt.Rows.Count+1,dt.Columns.Count);
            table.TableFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
            #region 测试
            //添加第1行
            TableRow rowhead = table.Rows[0];
            rowhead.IsHeader = true;
            rowhead.Height = 20;
            rowhead.HeightType = TableRowHeightType.Auto;
            rowhead.RowFormat.BackColor = Color.DarkGreen;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                rowhead.Cells[i].Width = 55;
                rowhead.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                rowhead.Height = 40;
                rowhead.HeightType = TableRowHeightType.Auto;
                 p = rowhead.Cells[i].AddParagraph();
                AddTextRange(p, dt.Columns[i].ColumnName, (float)10.5, true, "宋体", Spire.Doc.Documents.HorizontalAlignment.Center);
            }
            decimal statistics = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                TableRow dataRow = table.Rows[i + 1];
                dataRow.RowFormat.BackColor = Color.Empty;
                for (int r = 0; r < dt.Columns.Count; r++)
                {
                    dataRow.Cells[r].Width = 55;
                    dataRow.Cells[r].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    TextRange tr = dataRow.Cells[r].AddParagraph().AppendText(dt.Rows[i][r].ToString());
                    tr.CharacterFormat.FontSize = 9;
                }
                statistics +=Convert.ToDecimal(dt.Rows[i]["总价"]);
            }
            var count = doc.Sections.Count;
            //合计
            table.AddRow(true,9);
            table.ApplyHorizontalMerge(dt.Rows.Count + 1, 0,dt.Columns.Count-3);
           p = table.Rows[dt.Rows.Count + 1].Cells[0].AddParagraph();
            AddTextRange(p, "含税总金额(含税13%)", 10, true, "黑体", Spire.Doc.Documents.HorizontalAlignment.Center);
            p = table.Rows[dt.Rows.Count + 1].Cells[dt.Columns.Count - 2].AddParagraph();
            AddTextRange(p, statistics.ToString(), 10, true, "黑体", Spire.Doc.Documents.HorizontalAlignment.Left);
            if (disconuts>0)
            {
                table.AddRow(true, 9);
                table.ApplyHorizontalMerge(dt.Rows.Count + 2, 0, dt.Columns.Count - 3);
                p = table.Rows[dt.Rows.Count + 2].Cells[0].AddParagraph();
                AddTextRange(p, "最终优惠", 10, true, "黑体", Spire.Doc.Documents.HorizontalAlignment.Center);
                p = table.Rows[dt.Rows.Count + 2].Cells[dt.Columns.Count - 2].AddParagraph();
                AddTextRange(p, (statistics-disconuts).ToString(), 10, true, "黑体", Spire.Doc.Documents.HorizontalAlignment.Left);
            }
            #endregion

        }
        private void AddTextRange(Paragraph pragraph, string word, float fontSize, bool isBold, string fontName, Spire.Doc.Documents.HorizontalAlignment alignType)
        {
            TextRange textRange = pragraph.AppendText(word);
            textRange.CharacterFormat.FontSize = fontSize;
            textRange.CharacterFormat.Bold = isBold;
            textRange.CharacterFormat.FontName = fontName;
            pragraph.Format.HorizontalAlignment = alignType;
        }
    }
}
