using DocumentFormat.OpenXml.Packaging;
using OutPutExcel.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutPutExcel
{
    public class ExcelCommand
    {
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="filePath">
        /// The file path.
        /// </param>
        /// <param name="fileTemplatePath">
        /// The file template path.
        /// </param>
        /// <exception cref="Exception">
        /// </exception>
        public void ExcelOut(string filePath, string fileTemplatePath)
        {
            try
            {
                System.IO.File.Copy(fileTemplatePath, filePath);
            }
            catch (Exception ex)
            {
                throw new Exception("复制Excel文件出错" + ex.Message);
            }

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                var sheetData = document.GetFirstSheetData();
                OpenXmlHelper.CellStyleIndex = 1;

                var data = GetData();

                ////写标题相关信息

                const string C = "C", D = "D", A = "A", B = "B", E = "E";

                const int StartRowIndex =2, Len = 10;



                var index = 0; 
                foreach (var item in data)
                {
                    var rowIndex = StartRowIndex + index;

                    // 员工信息
                    sheetData.SetCellValue(A + rowIndex, rowIndex);
                    sheetData.SetCellValue(D + rowIndex, item.Date);
                    sheetData.SetCellValue(B + rowIndex, item.Name);

                    // 部门信息
                    sheetData.SetCellValue(C + rowIndex, item.BuMen);

                    // 备注
                    sheetData.SetCellValue(E + rowIndex, item.Remark);
                    index++;
                }

                //for (var i = 0; i < Len; i++)
                //{
                //    var rowIndex = StartRowIndex + i;

                //    // 员工信息
                //    sheetData.SetCellValue(A + rowIndex,  rowIndex);
                //    sheetData.SetCellValue(D + rowIndex, DateTime.Now.AddYears(-30).AddYears(new Random().Next(1, 30)));
                //    sheetData.SetCellValue(B + rowIndex, "姓名" + rowIndex);

                //    // 部门信息
                //    sheetData.SetCellValue(C + rowIndex, "部门" + rowIndex);

                //    // 备注
                //    sheetData.SetCellValue(E + rowIndex, "备注:" + rowIndex);
                //}

                // var str = OpenXmlHelper.ValidateDocument(document);验证生成的Excel
            }
        }
    

        public List<UserModel> GetData()
        {
            List<UserModel> users = new List<UserModel>();

            string[] bumens = new string[] { "技术部","市场部","财务部","人事部"};
            string[] remarks = new string[] { "优秀","一般","良好","差"};
            DateTime[] dates = new DateTime[] { DateTime.Now, DateTime.Now.AddDays(10), DateTime.Now.AddDays(20), DateTime.Now.AddDays(30) };
            var r = new Random();
            for (int i = 1; i <= 5000; i++)
            {
                var num = r.Next(0, 4);
                var user = new UserModel()
                {
                    ID = i,
                    Name = "用户"+i.ToString(),
                    Date = dates[num],
                    BuMen = bumens[num],
                    Remark = remarks[num],
                };
                users.Add(user);
            }
            return users;
        }
    }
}
